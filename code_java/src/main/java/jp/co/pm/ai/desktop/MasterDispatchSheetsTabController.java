package jp.co.pm.ai.desktop;

import java.nio.file.Path;
import java.time.OffsetDateTime;
import java.time.ZoneId;
import java.time.temporal.ChronoUnit;
import java.util.BitSet;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;

import javafx.application.Platform;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.geometry.Insets;
import javafx.scene.Node;
import javafx.scene.Parent;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Label;
import javafx.scene.control.TabPane;
import javafx.scene.control.TablePosition;
import javafx.scene.control.TextArea;
import javafx.scene.control.Tooltip;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.StackPane;
import javafx.stage.Window;

import org.controlsfx.control.spreadsheet.Grid;
import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.GridChange;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.io.MasterDispatchSheetsDocument;
import jp.co.pm.ai.desktop.io.MasterDispatchSheetsSaveWriter;
import jp.co.pm.ai.desktop.io.MasterDispatchSheetsSeeder;
import jp.co.pm.ai.desktop.io.MasterTeamCombinationTableReader;
import jp.co.pm.ai.desktop.ui.AttendanceGridLoadingOverlay;
import jp.co.pm.ai.desktop.ui.ButtonAttentionGlow;
import jp.co.pm.ai.desktop.ui.ColumnVisibilityDialog;
import jp.co.pm.ai.desktop.ui.ColumnVisibilitySupport;
import jp.co.pm.ai.desktop.ui.FourDigitConfirmationDialog;
import jp.co.pm.ai.desktop.ui.LabeledTextFillSupport;
import jp.co.pm.ai.desktop.ui.MasterDispatchCombinationRowDialog;
import jp.co.pm.ai.desktop.ui.MasterDispatchEquipmentColumnDialog;
import jp.co.pm.ai.desktop.ui.MasterDispatchSheetEditRules;
import jp.co.pm.ai.desktop.ui.MasterDispatchSheetGridSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetMultiColumnFilterCoordinator;
import jp.co.pm.ai.desktop.ui.SpreadsheetTabularSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetThemeBridge;

/**
 * メインタブ MASTER。現在工場の skills / need / speed / 組み合わせ表を JSON で表示・編集する。
 */
public final class MasterDispatchSheetsTabController {

    private static final ZoneId TOKYO = ZoneId.of("Asia/Tokyo");

    @FXML private Label statusLabel;
    @FXML private Label filterHintLabel;
    @FXML private TabPane innerTabPane;
    @FXML private StackPane skillsHost;
    @FXML private StackPane needHost;
    @FXML private StackPane speedHost;
    @FXML private StackPane comboHost;
    @FXML private Button saveButton;
    @FXML private Button addEquipmentColumnButton;
    @FXML private Button addMissingEquipmentButton;
    @FXML private Button addCombinationRowButton;
    @FXML private Button deleteCombinationRowButton;

    private final SpreadsheetView skillsView = new SpreadsheetView();
    private final SpreadsheetView needView = new SpreadsheetView();
    private final SpreadsheetView speedView = new SpreadsheetView();
    private final SpreadsheetView comboView = new SpreadsheetView();

    private MainShellController shell;
    private MasterDispatchSheetsDocument document = MasterDispatchSheetsDocument.empty("");
    private Path loadedJsonPath;
    /** 4タブ共通。工程+機械の正規化キー。空なら空き列以外をすべて表示。 */
    private final Set<String> equipmentFocusKeys = new LinkedHashSet<>();
    private ButtonAttentionGlow missingEquipmentGlow;
    private ButtonAttentionGlow saveButtonGlow;
    private int missingGlowEpoch;
    private int loadEpoch;
    private final AttendanceGridLoadingOverlay loadingOverlay =
            new AttendanceGridLoadingOverlay("pm-master-dispatch-grid-loading-overlay");
    private boolean suppressGridDirty;
    private final EventHandler<GridChange> gridDirtyHandler =
            e -> {
                if (!suppressGridDirty) {
                    applySaveDirtyState(true);
                }
            };

    @FXML
    private void initialize() {
        install(skillsHost, skillsView, 1);
        install(needHost, needView, 3);
        install(speedHost, speedView, 3);
        install(comboHost, comboView, 4);
        installLoadingOverlay();
        if (innerTabPane != null) {
            innerTabPane
                    .getSelectionModel()
                    .selectedIndexProperty()
                    .addListener(
                            (obs, o, n) -> {
                                applyAllEquipmentVisibility();
                                updateToolbarForInnerTab(
                                        n == null ? 0 : n.intValue());
                            });
            updateToolbarForInnerTab(innerTabPane.getSelectionModel().getSelectedIndex());
        }
        if (saveButton != null && saveButtonGlow == null) {
            saveButtonGlow = new ButtonAttentionGlow(saveButton);
        }
        applyDocument(document);
        applySaveDirtyState(false);
        scheduleEnsureReadableChromeTextColors();
    }

    private void installLoadingOverlay() {
        if (innerTabPane == null || !(innerTabPane.getParent() instanceof BorderPane bp)) {
            return;
        }
        StackPane stack = new StackPane();
        bp.setCenter(null);
        stack.getChildren().addAll(innerTabPane, loadingOverlay);
        bp.setCenter(stack);
    }

    private static void install(StackPane host, SpreadsheetView view, int leadingCols) {
        SpreadsheetThemeBridge.install(view);
        SpreadsheetTabularSupport.installPmAiReadableSpreadsheetChrome(view);
        if (!view.getStyleClass().contains("pm-ai-master-dispatch-spreadsheet")) {
            view.getStyleClass().add("pm-ai-master-dispatch-spreadsheet");
        }
        view.setEditable(true);
        StackPane.setMargin(view, new Insets(0));
        host.getChildren().setAll(view);
        MasterDispatchSheetGridSupport.installSingleClickListEditing(view);
        SpreadsheetTabularSupport.installSpreadsheetClickSelectionAlign(view);
        SpreadsheetTabularSupport.installSpreadsheetChromeRelayoutDebouncerForHost(
                host, () -> leadingCols, () -> view.getEditingCell() != null);
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
    }

    void onMainShellTabSelected() {
        reloadFromCurrentFactory(false);
        scheduleEnsureReadableChromeTextColors();
        if (shell != null) {
            Platform.runLater(
                    () -> {
                        shell.refreshMainShellTabHeaderChromeFromStoredColors();
                        ensureReadableChromeTextColors();
                    });
        }
    }

    /**
     * ミッドナイトブルー等で LabeledText が黒のまま残るのを、インライン {@code -fx-fill} で潰す。
     * Skin 未準備時の取りこぼし対策で runLater を重ねる。
     */
    private void scheduleEnsureReadableChromeTextColors() {
        ensureReadableChromeTextColors();
        Platform.runLater(this::ensureReadableChromeTextColors);
        Platform.runLater(
                () -> Platform.runLater(this::ensureReadableChromeTextColors));
    }

    private void ensureReadableChromeTextColors() {
        String mid = LabeledTextFillSupport.THEME_MID;
        LabeledTextFillSupport.applyToTabPaneHeaders(innerTabPane, mid);
        Parent root = findMasterDispatchRoot();
        if (root == null) {
            return;
        }
        for (Node node : root.lookupAll(".button")) {
            if (!(node instanceof Button button)) {
                continue;
            }
            boolean attention =
                    button.getStyleClass().contains("pm-aladdin-entry-export-attention");
            LabeledTextFillSupport.applyToButton(button, attention ? "#e0f2fe" : mid);
        }
    }

    private Parent findMasterDispatchRoot() {
        Node n = innerTabPane != null ? innerTabPane : saveButton;
        while (n != null) {
            if (n.getStyleClass().contains("pm-master-dispatch-sheets-tab")
                    && n instanceof Parent parent) {
                return parent;
            }
            n = n.getParent();
        }
        return null;
    }

    void reloadFromCurrentFactory(boolean reimport) {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        FactorySite site = AppPaths.currentDispatchFactorySite(ui);
        Path json = AppPaths.masterDispatchSheetsJsonPath(ui);
        Path source = AppPaths.masterDispatchSheetsSourceWorkbookPath(ui);
        boolean factoryChanged =
                document == null
                        || site == null
                        || !site.name().equals(document.factorySite())
                        || loadedJsonPath == null
                        || !loadedJsonPath.equals(json);
        if (!reimport && !factoryChanged && loadedJsonPath != null) {
            updateFilterHint();
            refreshMissingEquipmentAttention();
            scheduleEnsureReadableChromeTextColors();
            return;
        }
        if (reimport || factoryChanged) {
            equipmentFocusKeys.clear();
        }
        int epoch = ++loadEpoch;
        String loadingMessage =
                reimport ? "共有フォルダから再取込中" : "配台マスタを読み込み中";
        setLoadingVisible(true, loadingMessage);
        if (statusLabel != null) {
            statusLabel.setText(loadingMessage + "…");
        }
        final FactorySite siteRef = site;
        final Path jsonRef = json;
        final Path sourceRef = source;
        final boolean reimportRef = reimport;
        Thread worker =
                new Thread(
                        () -> {
                            try {
                                MasterDispatchSheetsSeeder.Result result =
                                        MasterDispatchSheetsSeeder.loadOrImport(
                                                jsonRef, sourceRef, (siteRef != null ? siteRef.name() : ""), reimportRef);
                                BuiltGrids built = buildGridsOffFx(result.document());
                                Platform.runLater(
                                        () -> {
                                            if (epoch != loadEpoch) {
                                                return;
                                            }
                                            document = result.document();
                                            loadedJsonPath = jsonRef;
                                            applyBuiltGridsProgressively(
                                                    built,
                                                    () -> {
                                                        if (epoch != loadEpoch) {
                                                            return;
                                                        }
                                                        statusLabel.setText(
                                                                statusText(
                                                                        siteRef,
                                                                        jsonRef,
                                                                        sourceRef,
                                                                        result));
                                                        applySaveDirtyState(false);
                                                        setLoadingVisible(false, null);
                                                        scheduleEnsureReadableChromeTextColors();
                                                    });
                                        });
                            } catch (Exception e) {
                                Platform.runLater(
                                        () -> {
                                            if (epoch != loadEpoch) {
                                                return;
                                            }
                                            statusLabel.setText(
                                                    "読込に失敗しました: " + e.getMessage());
                                            setLoadingVisible(false, null);
                                        });
                            }
                        },
                        "master-dispatch-reload");
        worker.setDaemon(true);
        worker.start();
    }

    private void setLoadingVisible(boolean visible, String message) {
        if (visible) {
            loadingOverlay.setLoading(true, message);
        } else {
            loadingOverlay.setLoading(false);
        }
    }

    private record BuiltGrids(
            GridBase skills,
            GridBase need,
            GridBase speed,
            GridBase combo,
            List<List<String>> skillsRows,
            List<List<String>> needRows,
            List<List<String>> speedRows,
            List<List<String>> comboRows) {}

    private static BuiltGrids buildGridsOffFx(MasterDispatchSheetsDocument doc) {
        MasterDispatchSheetsDocument d =
                doc != null ? doc : MasterDispatchSheetsDocument.empty("");
        List<List<String>> skillsRows = d.sheet("skills").rows();
        List<List<String>> needRows = d.sheet("need").rows();
        List<List<String>> speedRows = d.sheet("speed").rows();
        List<List<String>> comboRows =
                MasterDispatchSheetEditRules.ensureCombinationMetaColumns(
                        d.sheet("teamCombinations").rows());
        return new BuiltGrids(
                MasterDispatchSheetGridSupport.buildEditable(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS, skillsRows),
                MasterDispatchSheetGridSupport.buildEditable(
                        MasterDispatchSheetEditRules.SheetKind.NEED, needRows),
                MasterDispatchSheetGridSupport.buildEditable(
                        MasterDispatchSheetEditRules.SheetKind.SPEED, speedRows),
                MasterDispatchSheetGridSupport.buildEditable(
                        MasterDispatchSheetEditRules.SheetKind.COMBINATIONS,
                        comboRows,
                        skillsRows),
                skillsRows,
                needRows,
                speedRows,
                comboRows);
    }

    /** 選択中の内タブを先に載せ、残りはパルスごとに 1 枚ずつ載せて UI を固まらせない。 */
    private void applyBuiltGridsProgressively(BuiltGrids built, Runnable onComplete) {
        if (built == null) {
            if (onComplete != null) {
                onComplete.run();
            }
            return;
        }
        suppressGridDirty = true;
        int selected =
                innerTabPane != null ? innerTabPane.getSelectionModel().getSelectedIndex() : 0;
        if (selected < 0) {
            selected = 0;
        }
        int[] order = new int[4];
        order[0] = selected;
        int w = 1;
        for (int i = 0; i < 4; i++) {
            if (i != selected) {
                order[w++] = i;
            }
        }
        attachBuiltGrid(order[0], built);
        applyOneEquipmentVisibility(order[0], built);
        updateToolbarForInnerTab(selected);
        attachBuiltGridsRemaining(built, order, 1, onComplete);
    }

    private void attachBuiltGridsRemaining(
            BuiltGrids built, int[] order, int nextIndex, Runnable onComplete) {
        if (nextIndex >= order.length) {
            refreshMissingEquipmentAttention();
            Platform.runLater(
                    () -> {
                        suppressGridDirty = false;
                        if (onComplete != null) {
                            onComplete.run();
                        }
                    });
            return;
        }
        Platform.runLater(
                () -> {
                    attachBuiltGrid(order[nextIndex], built);
                    applyOneEquipmentVisibility(order[nextIndex], built);
                    attachBuiltGridsRemaining(built, order, nextIndex + 1, onComplete);
                });
    }

    private void applyOneEquipmentVisibility(int index, BuiltGrids built) {
        switch (index) {
            case 0 ->
                    applyEquipmentColumnVisibility(
                            skillsView,
                            MasterDispatchSheetEditRules.SheetKind.SKILLS,
                            built.skillsRows(),
                            1,
                            equipmentFocusKeys);
            case 1 ->
                    applyEquipmentColumnVisibility(
                            needView,
                            MasterDispatchSheetEditRules.SheetKind.NEED,
                            built.needRows(),
                            3,
                            equipmentFocusKeys);
            case 2 ->
                    applyEquipmentColumnVisibility(
                            speedView,
                            MasterDispatchSheetEditRules.SheetKind.SPEED,
                            built.speedRows(),
                            3,
                            equipmentFocusKeys);
            case 3 -> applyCombinationRowVisibility();
            default -> {
                /* no-op */
            }
        }
        updateFilterHint();
    }

    private void attachBuiltGrid(int index, BuiltGrids built) {
        switch (index) {
            case 0 ->
                    attachPrebuiltGrid(
                            skillsView,
                            MasterDispatchSheetEditRules.SheetKind.SKILLS,
                            built.skills(),
                            built.skillsRows(),
                            MasterDispatchSheetEditRules.frozenTitleRowCount(
                                    MasterDispatchSheetEditRules.SheetKind.SKILLS),
                            1);
            case 1 ->
                    attachPrebuiltGrid(
                            needView,
                            MasterDispatchSheetEditRules.SheetKind.NEED,
                            built.need(),
                            built.needRows(),
                            MasterDispatchSheetEditRules.frozenTitleRowCount(
                                    MasterDispatchSheetEditRules.SheetKind.NEED),
                            3);
            case 2 ->
                    attachPrebuiltGrid(
                            speedView,
                            MasterDispatchSheetEditRules.SheetKind.SPEED,
                            built.speed(),
                            built.speedRows(),
                            MasterDispatchSheetEditRules.frozenTitleRowCount(
                                    MasterDispatchSheetEditRules.SheetKind.SPEED),
                            3);
            case 3 ->
                    attachPrebuiltGrid(
                            comboView,
                            MasterDispatchSheetEditRules.SheetKind.COMBINATIONS,
                            built.combo(),
                            built.comboRows(),
                            0,
                            4);
            default -> {
                /* no-op */
            }
        }
    }

    @FXML
    private void onSaveAction() {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        FactorySite site = AppPaths.currentDispatchFactorySite(ui);
        Path json = loadedJsonPath != null ? loadedJsonPath : AppPaths.masterDispatchSheetsJsonPath(ui);
        Path source = AppPaths.masterDispatchSheetsSourceWorkbookPath(ui);
        List<List<String>> skillsRows =
                MasterDispatchSheetGridSupport.extract(
                        skillsView, MasterDispatchSheetEditRules.SheetKind.SKILLS);
        List<List<String>> needRows =
                MasterDispatchSheetGridSupport.extract(
                        needView, MasterDispatchSheetEditRules.SheetKind.NEED);
        List<List<String>> speedRows =
                MasterDispatchSheetGridSupport.extract(
                        speedView, MasterDispatchSheetEditRules.SheetKind.SPEED);
        List<List<String>> comboRows =
                MasterDispatchSheetGridSupport.extract(
                        comboView, MasterDispatchSheetEditRules.SheetKind.COMBINATIONS);
        List<String> errors = new java.util.ArrayList<>();
        errors.addAll(
                MasterDispatchSheetEditRules.validateForSave(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS, skillsRows));
        errors.addAll(
                MasterDispatchSheetEditRules.validateForSave(
                        MasterDispatchSheetEditRules.SheetKind.NEED, needRows));
        errors.addAll(
                MasterDispatchSheetEditRules.validateForSave(
                        MasterDispatchSheetEditRules.SheetKind.SPEED, speedRows));
        errors.addAll(
                MasterDispatchSheetEditRules.validateForSave(
                        MasterDispatchSheetEditRules.SheetKind.COMBINATIONS, comboRows));
        if (!errors.isEmpty()) {
            Alert alert = new Alert(AlertType.ERROR);
            alert.setTitle("配台マスタ保存");
            alert.setHeaderText(
                    "入力内容を直してから保存してください（"
                            + errors.size()
                            + " 件）。");
            int shown = Math.min(12, errors.size());
            String body = String.join("\n", errors.subList(0, shown));
            if (errors.size() > shown) {
                body = body + "\n…ほか " + (errors.size() - shown) + " 件";
            }
            TextArea area = new TextArea(body);
            area.setEditable(false);
            area.setWrapText(true);
            area.setPrefRowCount(Math.min(12, Math.max(4, shown)));
            area.setPrefWidth(520);
            alert.getDialogPane().setContent(area);
            Window w = statusLabel != null && statusLabel.getScene() != null ? statusLabel.getScene().getWindow() : null;
            if (w != null) {
                alert.initOwner(w);
            }
            alert.showAndWait();
            statusLabel.setText("保存を中止しました。検証エラーが " + errors.size() + " 件あります。");
            return;
        }
        Window owner = shell.primaryStageForDialogs();
        if (owner == null && statusLabel != null && statusLabel.getScene() != null) {
            owner = statusLabel.getScene().getWindow();
        }
        if (!FourDigitConfirmationDialog.confirm(
                owner,
                "配台マスタ保存",
                "編集内容を JSON に保存し、master.xlsm へ書き戻します。\n"
                        + "保存前に JSON と master.xlsm の世代バックアップを取ります。\n"
                        + "誤操作防止のため、表示されるランダムな4桁の数字を入力してください。",
                "保存")) {
            statusLabel.setText("保存を中止しました。");
            return;
        }
        LinkedHashMap<String, MasterDispatchSheetsDocument.SheetGrid> sheets = new LinkedHashMap<>();
        sheets.put(
                MasterDispatchSheetsDocument.KEY_SKILLS,
                new MasterDispatchSheetsDocument.SheetGrid("skills", skillsRows));
        sheets.put(
                MasterDispatchSheetsDocument.KEY_NEED,
                new MasterDispatchSheetsDocument.SheetGrid("need", needRows));
        sheets.put(
                MasterDispatchSheetsDocument.KEY_SPEED,
                new MasterDispatchSheetsDocument.SheetGrid("speed", speedRows));
        sheets.put(
                MasterDispatchSheetsDocument.KEY_TEAM_COMBINATIONS,
                new MasterDispatchSheetsDocument.SheetGrid("組み合わせ表", comboRows));
        document =
                new MasterDispatchSheetsDocument(
                        MasterDispatchSheetsDocument.SCHEMA_VERSION,
                        site.name(),
                        source.toString(),
                        OffsetDateTime.now(TOKYO).truncatedTo(ChronoUnit.SECONDS).toString(),
                        sheets);
        try {
            MasterDispatchSheetsSaveWriter.Result saved =
                    MasterDispatchSheetsSaveWriter.save(json, source, document, ui);
            loadedJsonPath = saved.jsonPath();
            String jsonBackup =
                    saved.jsonBackup() != null ? saved.jsonBackup().toString() : "（JSON 正本なし）";
            String xlsmBackup =
                    saved.workbookBackup() != null
                            ? saved.workbookBackup().toString()
                            : "（ブック正本なし）";
            statusLabel.setText(
                    "保存しました。JSON: "
                            + saved.jsonPath()
                            + "\nmaster: "
                            + saved.workbookPath()
                            + "\n世代バックアップ JSON: "
                            + jsonBackup
                            + "\n世代バックアップ master: "
                            + xlsmBackup);
            refreshMissingEquipmentAttention();
            applySaveDirtyState(false);
        } catch (Exception e) {
            statusLabel.setText("保存に失敗しました: " + e.getMessage());
        }
    }

    @FXML
    private void onReimportAction() {
        Alert confirm = new Alert(AlertType.CONFIRMATION);
        confirm.setTitle("共有フォルダから再取込");
        confirm.setHeaderText("現在工場の JSON を Excel で上書きします。");
        confirm.setContentText("手編集は失われます。よろしいですか。");
        Window w = statusLabel != null && statusLabel.getScene() != null ? statusLabel.getScene().getWindow() : null;
        if (w != null) {
            confirm.initOwner(w);
        }
        Optional<ButtonType> ans = confirm.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            return;
        }
        reloadFromCurrentFactory(true);
    }

    @FXML
    private void onAddEquipmentColumnAction() {
        Window owner = dialogOwner();
        Optional<MasterDispatchEquipmentColumnDialog.Result> ans =
                MasterDispatchEquipmentColumnDialog.prompt(owner);
        if (ans.isEmpty()) {
            return;
        }
        MasterDispatchEquipmentColumnDialog.Result r = ans.get();
        int added = applyEquipmentColumn(r.process(), r.machine());
        if (added <= 0) {
            showInfo(
                    "設備列を追加",
                    "既に同じ工程名+機械名の列があります。",
                    r.process() + " × " + r.machine());
            statusLabel.setText("既に同じ工程名+機械名の列があります: " + r.process() + " × " + r.machine());
            return;
        }
        statusLabel.setText(
                "列を追加しました: "
                        + r.process()
                        + " × "
                        + r.machine()
                        + "。資格の OP/AS と必要人数を入れてから保存してください。");
    }

    @FXML
    private void onAddMissingEquipmentColumnsAction() {
        addMissingEquipmentColumnsFromPlan(true);
    }

    /**
     * 段階1/2 の未登録ダイアログから呼ぶ。計画タスクにあって skills に無い工程+機械列を追加する。
     *
     * @return 新規に足した列数
     */
    int addMissingEquipmentColumnsFromPlan(boolean showAlerts) {
        if (shell == null) {
            return 0;
        }
        PlanTasksMissingSkillsColumnPrompt.PromptBundle bundle;
        try {
            bundle = PlanTasksMissingSkillsColumnPrompt.collectMissingPairs(shell.snapshotUiEnv());
        } catch (Exception e) {
            if (showAlerts) {
                statusLabel.setText("未登録の工程+機械を確認できませんでした: " + e.getMessage());
            }
            return 0;
        }
        if (bundle.empty()) {
            if (showAlerts) {
                Alert alert = new Alert(AlertType.INFORMATION);
                alert.setTitle("未登録の工程+機械");
                alert.setHeaderText("追加する列はありません。");
                alert.setContentText("計画タスクの工程+機械は、すべて skills シートに列があります。");
                Window w = dialogOwner();
                if (w != null) {
                    alert.initOwner(w);
                }
                alert.showAndWait();
            }
            return 0;
        }
        equipmentFocusKeys.clear();
        equipmentFocusKeys.addAll(focusKeysForMissingPairs(bundle.pairs()));
        int added = 0;
        for (PlanTasksMissingSkillsColumnPrompt.MissingPair pair : bundle.pairs()) {
            added += applyEquipmentColumn(pair.process(), pair.machine());
        }
        focusSkillsTab();
        statusLabel.setText(
                "未登録の工程+機械を "
                        + added
                        + " 列追加しました。追加した列だけ表示しています。OP/AS を設定し、保存してください。「すべて表示」で全列に戻せます。");
        refreshMissingEquipmentAttention();
        return added;
    }

    int addMissingEquipmentColumns(List<PlanTasksMissingSkillsColumnPrompt.MissingPair> pairs) {
        if (pairs == null || pairs.isEmpty()) {
            return 0;
        }
        equipmentFocusKeys.clear();
        equipmentFocusKeys.addAll(focusKeysForMissingPairs(pairs));
        int added = 0;
        for (PlanTasksMissingSkillsColumnPrompt.MissingPair pair : pairs) {
            added += applyEquipmentColumn(pair.process(), pair.machine());
        }
        focusSkillsTab();
        statusLabel.setText(
                "未登録の工程+機械を "
                        + added
                        + " 列追加しました。追加した列だけ表示しています。OP/AS を設定し、保存してください。「すべて表示」で全列に戻せます。");
        refreshMissingEquipmentAttention();
        return added;
    }

    static Set<String> focusKeysForMissingPairs(
            List<PlanTasksMissingSkillsColumnPrompt.MissingPair> pairs) {
        LinkedHashSet<String> keys = new LinkedHashSet<>();
        if (pairs == null) {
            return keys;
        }
        for (PlanTasksMissingSkillsColumnPrompt.MissingPair pair : pairs) {
            if (pair == null) {
                continue;
            }
            String key =
                    MasterTeamCombinationTableReader.normalizedComboKey(
                            pair.process(), pair.machine());
            if (!key.isEmpty()) {
                keys.add(key);
            }
        }
        return keys;
    }

    @FXML
    private void onChooseVisibleEquipmentAction() {
        if (skillsView.getGrid() == null) {
            return;
        }
        int colCount = skillsView.getGrid().getColumnCount();
        List<String> titles =
                MasterDispatchSheetEditRules.columnTitles(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS,
                        document.sheet("skills").rows(),
                        colCount);
        List<String> labels = MasterDispatchSheetEditRules.dialogColumnLabels(titles);
        boolean[] vis = MasterDispatchSheetEditRules.visibilityMask(titles, 1, equipmentFocusKeys);
        boolean[] mandatory = MasterDispatchSheetEditRules.mandatoryLeadingMask(titles.size(), 1);
        ColumnVisibilityDialog.show(dialogOwner(), labels, vis, mandatory)
                .ifPresent(
                        arr -> {
                            boolean[] merged =
                                    ColumnVisibilitySupport.mergeMandatoryIntoVisibility(
                                            arr, mandatory);
                            equipmentFocusKeys.clear();
                            equipmentFocusKeys.addAll(
                                    MasterDispatchSheetEditRules.focusKeysFromVisibility(
                                            titles, 1, merged));
                            applyAllEquipmentVisibility();
                            statusLabel.setText(
                                    "表示する設備を資格・必要人数・加工速度・組み合わせ表に適用しました。保存時に非表示のデータは消えません。");
                        });
    }

    @FXML
    private void onShowAllEquipmentColumnsAction() {
        equipmentFocusKeys.clear();
        applyAllEquipmentVisibility();
        statusLabel.setText("すべての設備を表示しています（空き列は隠しています）。");
    }

    @FXML
    private void onAddCombinationRowAction() {
        List<List<String>> skills =
                MasterDispatchSheetGridSupport.extract(
                        skillsView, MasterDispatchSheetEditRules.SheetKind.SKILLS);
        List<List<String>> combo =
                MasterDispatchSheetGridSupport.extract(
                        comboView, MasterDispatchSheetEditRules.SheetKind.COMBINATIONS);
        List<String[]> pairs = MasterDispatchSheetEditRules.skillsEquipmentPairs(skills);
        Optional<MasterDispatchCombinationRowDialog.Result> ans =
                MasterDispatchCombinationRowDialog.prompt(dialogOwner(), pairs);
        if (ans.isEmpty()) {
            return;
        }
        MasterDispatchCombinationRowDialog.Result r = ans.get();
        if (MasterDispatchSheetEditRules.containsCombinationEquipment(combo, r.process(), r.machine())) {
            showInfo(
                    "組み合わせ行を追加",
                    "同じ工程名+機械名の行が既にあります。",
                    r.process() + " × " + r.machine());
            statusLabel.setText(
                    "同じ工程名+機械名の行が既にあります: " + r.process() + " × " + r.machine());
            focusComboTab();
            return;
        }
        List<List<String>> added =
                MasterDispatchSheetEditRules.addCombinationRow(
                        combo,
                        r.process(),
                        r.machine(),
                        r.autoFillSkillMembers() ? skills : null);
        replaceCurrentSheetsKeepingNeedSpeed(skills, added);
        focusComboTab();
        if (r.autoFillSkillMembers()) {
            int filled =
                    MasterDispatchSheetEditRules.skilledMembersForEquipment(
                                    skills, r.process(), r.machine())
                            .size();
            int memberCols = 0;
            if (!added.isEmpty()) {
                List<String> header = added.get(0);
                for (int c = 0; c < header.size(); c++) {
                    if (MasterDispatchSheetEditRules.isCombinationMemberColumn(header, c)) {
                        memberCols++;
                    }
                }
            }
            int placed = Math.min(filled, memberCols);
            if (placed > 0) {
                statusLabel.setText(
                        "組み合わせ行を追加しました: "
                                + r.process()
                                + " × "
                                + r.machine()
                                + "。スキルメンバーを "
                                + placed
                                + " 人入れました。追加行は色が違います。");
            } else {
                statusLabel.setText(
                        "組み合わせ行を追加しました: "
                                + r.process()
                                + " × "
                                + r.machine()
                                + "。スキル該当者がいないためメンバーは空です。追加行は色が違います。");
            }
        } else {
            statusLabel.setText(
                    "組み合わせ行を追加しました: "
                            + r.process()
                            + " × "
                            + r.machine()
                            + "。追加行は色が違います。メンバーは OP/AS 付きで選んでください。");
        }
    }

    @FXML
    private void onDeleteCombinationRowAction() {
        if (!isComboInnerTabSelected()) {
            statusLabel.setText("組み合わせ表を開いてから、削除する行を選んでください。");
            return;
        }
        if (comboView.getSelectionModel() == null
                || comboView.getSelectionModel().getSelectedCells().isEmpty()) {
            statusLabel.setText("削除する組み合わせ表の行を選んでください。");
            return;
        }
        int first = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        Set<Integer> modelRows = new LinkedHashSet<>();
        for (TablePosition<?, ?> pos : comboView.getSelectionModel().getSelectedCells()) {
            int viewRow = pos.getRow();
            if (viewRow < 0) {
                continue;
            }
            modelRows.add(comboView.getModelRow(viewRow));
        }
        Set<Integer> originalIndexes = combinationDocumentIndexesFromModelRows(modelRows, first);
        List<List<String>> combo =
                MasterDispatchSheetGridSupport.extract(
                        comboView, MasterDispatchSheetEditRules.SheetKind.COMBINATIONS);
        List<String> targets = combinationDeleteLabels(combo, originalIndexes);
        if (targets.isEmpty()) {
            statusLabel.setText("ロック中の行は削除できません。編集ロックを外してから削除してください。");
            focusComboTab();
            return;
        }
        Alert confirm = new Alert(AlertType.CONFIRMATION);
        confirm.setTitle("選択行を削除");
        confirm.setHeaderText("組み合わせ表の行を削除します。ロック行は残します。");
        confirm.setContentText(String.join("\n", targets));
        Window owner = dialogOwner();
        if (owner != null) {
            confirm.initOwner(owner);
        }
        Optional<ButtonType> ans = confirm.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            return;
        }
        int before = combo.size();
        List<List<String>> after =
                MasterDispatchSheetEditRules.deleteCombinationRows(combo, originalIndexes);
        if (after.size() == before) {
            statusLabel.setText("ロック中の行は削除できません。編集ロックを外してから削除してください。");
            focusComboTab();
            return;
        }
        List<List<String>> skills =
                MasterDispatchSheetGridSupport.extract(
                        skillsView, MasterDispatchSheetEditRules.SheetKind.SKILLS);
        replaceCurrentSheetsKeepingNeedSpeed(skills, after);
        focusComboTab();
        statusLabel.setText("選択した組み合わせ行を削除しました（ロック行は残しています）。");
    }

    private void replaceCurrentSheetsKeepingNeedSpeed(
            List<List<String>> skills, List<List<String>> combo) {
        List<List<String>> need =
                MasterDispatchSheetGridSupport.extract(
                        needView, MasterDispatchSheetEditRules.SheetKind.NEED);
        List<List<String>> speed =
                MasterDispatchSheetGridSupport.extract(
                        speedView, MasterDispatchSheetEditRules.SheetKind.SPEED);
        replaceDocumentSheets(skills, need, speed, combo);
    }

    void focusComboTab() {
        if (innerTabPane != null && innerTabPane.getTabs().size() > 3) {
            innerTabPane.getSelectionModel().select(3);
        }
    }

    void focusSkillsTab() {
        if (innerTabPane != null && !innerTabPane.getTabs().isEmpty()) {
            innerTabPane.getSelectionModel().select(0);
        }
    }

    private int applyEquipmentColumn(String process, String machine) {
        List<List<String>> skills =
                MasterDispatchSheetGridSupport.extract(
                        skillsView, MasterDispatchSheetEditRules.SheetKind.SKILLS);
        List<List<String>> need =
                MasterDispatchSheetGridSupport.extract(
                        needView, MasterDispatchSheetEditRules.SheetKind.NEED);
        List<List<String>> speed =
                MasterDispatchSheetGridSupport.extract(
                        speedView, MasterDispatchSheetEditRules.SheetKind.SPEED);
        List<List<String>> combo =
                MasterDispatchSheetGridSupport.extract(
                        comboView, MasterDispatchSheetEditRules.SheetKind.COMBINATIONS);
        boolean existed =
                MasterDispatchSheetEditRules.containsEquipmentColumn(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS, skills, process, machine);
        if (!existed) {
            String key = MasterTeamCombinationTableReader.normalizedComboKey(process, machine);
            if (!key.isEmpty() && !equipmentFocusKeys.isEmpty()) {
                equipmentFocusKeys.add(key);
            }
        }
        skills =
                MasterDispatchSheetEditRules.addEquipmentColumn(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS, skills, process, machine);
        need =
                MasterDispatchSheetEditRules.addEquipmentColumn(
                        MasterDispatchSheetEditRules.SheetKind.NEED, need, process, machine);
        speed =
                MasterDispatchSheetEditRules.addEquipmentColumn(
                        MasterDispatchSheetEditRules.SheetKind.SPEED, speed, process, machine);
        replaceDocumentSheets(skills, need, speed, combo);
        return existed ? 0 : 1;
    }

    private void replaceDocumentSheets(
            List<List<String>> skills,
            List<List<String>> need,
            List<List<String>> speed,
            List<List<String>> combo) {
        String site = document != null ? document.factorySite() : "";
        String source = document != null ? document.sourceWorkbook() : "";
        String imported = document != null ? document.importedAt() : "";
        LinkedHashMap<String, MasterDispatchSheetsDocument.SheetGrid> sheets = new LinkedHashMap<>();
        sheets.put(
                MasterDispatchSheetsDocument.KEY_SKILLS,
                new MasterDispatchSheetsDocument.SheetGrid("skills", skills));
        sheets.put(
                MasterDispatchSheetsDocument.KEY_NEED,
                new MasterDispatchSheetsDocument.SheetGrid("need", need));
        sheets.put(
                MasterDispatchSheetsDocument.KEY_SPEED,
                new MasterDispatchSheetsDocument.SheetGrid("speed", speed));
        sheets.put(
                MasterDispatchSheetsDocument.KEY_TEAM_COMBINATIONS,
                new MasterDispatchSheetsDocument.SheetGrid("組み合わせ表", combo));
        document =
                new MasterDispatchSheetsDocument(
                        MasterDispatchSheetsDocument.SCHEMA_VERSION, site, source, imported, sheets);
        applyDocument(document);
        applySaveDirtyState(true);
    }

    private Window dialogOwner() {
        if (shell != null && shell.primaryStageForDialogs() != null) {
            return shell.primaryStageForDialogs();
        }
        if (statusLabel != null && statusLabel.getScene() != null) {
            return statusLabel.getScene().getWindow();
        }
        return null;
    }

    private void applyDocument(MasterDispatchSheetsDocument doc) {
        this.document = doc != null ? doc : MasterDispatchSheetsDocument.empty("");
        MasterDispatchSheetsDocument d = this.document;
        suppressGridDirty = true;
        try {
            attachGrid(
                    skillsView,
                    MasterDispatchSheetEditRules.SheetKind.SKILLS,
                    d.sheet("skills").rows(),
                    MasterDispatchSheetEditRules.frozenTitleRowCount(
                            MasterDispatchSheetEditRules.SheetKind.SKILLS),
                    1);
            attachGrid(
                    needView,
                    MasterDispatchSheetEditRules.SheetKind.NEED,
                    d.sheet("need").rows(),
                    MasterDispatchSheetEditRules.frozenTitleRowCount(
                            MasterDispatchSheetEditRules.SheetKind.NEED),
                    3);
            attachGrid(
                    speedView,
                    MasterDispatchSheetEditRules.SheetKind.SPEED,
                    d.sheet("speed").rows(),
                    MasterDispatchSheetEditRules.frozenTitleRowCount(
                            MasterDispatchSheetEditRules.SheetKind.SPEED),
                    3);
            attachGrid(
                    comboView,
                    MasterDispatchSheetEditRules.SheetKind.COMBINATIONS,
                    d.sheet("teamCombinations").rows(),
                    0,
                    4,
                    d.sheet("skills").rows());
        } finally {
            Platform.runLater(() -> Platform.runLater(() -> suppressGridDirty = false));
        }
        applyAllEquipmentVisibility();
        updateToolbarForInnerTab(
                innerTabPane != null ? innerTabPane.getSelectionModel().getSelectedIndex() : 0);
        refreshMissingEquipmentAttention();
    }

    private void applyAllEquipmentVisibility() {
        applyEquipmentColumnVisibility(
                skillsView,
                MasterDispatchSheetEditRules.SheetKind.SKILLS,
                document.sheet("skills").rows(),
                1,
                equipmentFocusKeys);
        applyEquipmentColumnVisibility(
                needView,
                MasterDispatchSheetEditRules.SheetKind.NEED,
                document.sheet("need").rows(),
                3,
                equipmentFocusKeys);
        applyEquipmentColumnVisibility(
                speedView,
                MasterDispatchSheetEditRules.SheetKind.SPEED,
                document.sheet("speed").rows(),
                3,
                equipmentFocusKeys);
        applyCombinationRowVisibility();
        updateFilterHint();
    }

    private void updateToolbarForInnerTab(int index) {
        boolean combo = index == 3;
        setToolbarManaged(addEquipmentColumnButton, !combo);
        setToolbarManaged(addMissingEquipmentButton, !combo);
        setToolbarManaged(addCombinationRowButton, combo);
        setToolbarManaged(deleteCombinationRowButton, combo);
        updateFilterHint();
        ensureReadableChromeTextColors();
    }

    private static void setToolbarManaged(Button button, boolean show) {
        if (button == null) {
            return;
        }
        button.setVisible(show);
        button.setManaged(show);
    }

    private boolean isComboInnerTabSelected() {
        return innerTabPane != null && innerTabPane.getSelectionModel().getSelectedIndex() == 3;
    }

    private void updateFilterHint() {
        if (filterHintLabel == null) {
            return;
        }
        if (equipmentFocusKeys.isEmpty()) {
            filterHintLabel.setText("すべての設備を表示しています（空き列は隠しています）。保存しても非表示データは消えません。");
            return;
        }
        boolean combo = isComboInnerTabSelected();
        filterHintLabel.setText(
                "設備絞り込み中（"
                        + equipmentFocusKeys.size()
                        + " 件）。"
                        + (combo ? "組み合わせ表は行を隠しています。" : "資格・必要人数・加工速度は列を隠しています。")
                        + " データは残ります。「すべて表示」で戻せます。");
    }

    private void applySaveDirtyState(boolean dirty) {
        if (saveButtonGlow == null && saveButton != null) {
            saveButtonGlow = new ButtonAttentionGlow(saveButton);
        }
        if (saveButtonGlow != null) {
            if (dirty) {
                saveButtonGlow.startIfIdle();
            } else {
                saveButtonGlow.stop();
            }
        }
        if (saveButton != null) {
            saveButton.setTooltip(
                    dirty
                            ? new Tooltip("未保存の変更があります。保存時は表示される4桁の数字を入力してください。")
                            : null);
        }
    }

    private void refreshMissingEquipmentAttention() {
        if (addMissingEquipmentButton == null || shell == null) {
            return;
        }
        if (missingEquipmentGlow == null) {
            missingEquipmentGlow = new ButtonAttentionGlow(addMissingEquipmentButton);
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        List<List<String>> skills = currentSkillsRowsForMissingCheck();
        int epoch = ++missingGlowEpoch;
        Thread worker =
                new Thread(
                        () -> {
                            boolean has = false;
                            int count = 0;
                            try {
                                var keys =
                                        PlanTasksMissingSkillsColumnPrompt.normalizedSkillsKeys(
                                                skills);
                                var bundle =
                                        PlanTasksMissingSkillsColumnPrompt
                                                .collectMissingAgainstSkillsKeys(ui, keys);
                                has = !bundle.empty();
                                count = has ? bundle.pairs().size() : 0;
                            } catch (Exception ignored) {
                                has = false;
                            }
                            boolean glow = has;
                            int n = count;
                            Platform.runLater(
                                    () -> {
                                        if (epoch != missingGlowEpoch) {
                                            return;
                                        }
                                        applyMissingEquipmentGlow(glow, n);
                                    });
                        },
                        "master-dispatch-missing-glow");
        worker.setDaemon(true);
        worker.start();
    }

    private List<List<String>> currentSkillsRowsForMissingCheck() {
        if (skillsView.getGrid() != null) {
            return MasterDispatchSheetGridSupport.extract(
                    skillsView, MasterDispatchSheetEditRules.SheetKind.SKILLS);
        }
        if (document != null) {
            return document.sheet("skills").rows();
        }
        return List.of();
    }

    private void applyMissingEquipmentGlow(boolean hasMissing, int count) {
        if (addMissingEquipmentButton == null) {
            return;
        }
        if (hasMissing) {
            missingEquipmentGlow.startIfIdle();
            addMissingEquipmentButton.setTooltip(
                    new Tooltip(
                            "計画タスクに未登録の工程+機械が "
                                    + count
                                    + " 件あります。押すと資格・必要人数・加工速度へ列を追加します。"));
        } else {
            ButtonAttentionGlow.stopAll(missingEquipmentGlow);
            addMissingEquipmentButton.setTooltip(
                    new Tooltip("計画タスクの工程+機械で、資格シートに無い列を追加します。"));
        }
        scheduleEnsureReadableChromeTextColors();
    }

    private void showInfo(String title, String header, String content) {
        Alert alert = new Alert(AlertType.INFORMATION);
        alert.setTitle(title);
        alert.setHeaderText(header);
        alert.setContentText(content);
        Window w = dialogOwner();
        if (w != null) {
            alert.initOwner(w);
        }
        alert.showAndWait();
    }

    /** SpreadsheetView の model 行（非表示行を含む格子行）から組み合わせ文書の行インデックスへ。 */
    static Set<Integer> combinationDocumentIndexesFromModelRows(
            Iterable<Integer> modelGridRows, int firstDataGridRow) {
        Set<Integer> out = new LinkedHashSet<>();
        if (modelGridRows == null) {
            return out;
        }
        for (Integer modelRow : modelGridRows) {
            if (modelRow == null) {
                continue;
            }
            int displayIndex = modelRow - firstDataGridRow;
            if (displayIndex >= 0) {
                out.add(displayIndex + 1);
            }
        }
        return out;
    }

    static List<String> combinationDeleteLabels(
            List<List<String>> combo, Set<Integer> originalIndexes) {
        List<String> labels = new java.util.ArrayList<>();
        if (combo == null || combo.isEmpty() || originalIndexes == null) {
            return labels;
        }
        List<String> header = combo.get(0);
        int procCol = MasterDispatchSheetEditRules.headerIndex(header, "工程名");
        int machCol = MasterDispatchSheetEditRules.headerIndex(header, "機械名");
        int lockCol =
                MasterDispatchSheetEditRules.headerIndex(
                        header, MasterDispatchSheetEditRules.COL_EDIT_LOCK);
        for (int idx : originalIndexes) {
            if (idx <= 0 || idx >= combo.size()) {
                continue;
            }
            List<String> row = combo.get(idx);
            if (row == null) {
                continue;
            }
            String lock =
                    lockCol >= 0 && lockCol < row.size() && row.get(lockCol) != null
                            ? row.get(lockCol).strip()
                            : "";
            if (MasterDispatchSheetEditRules.isCombinationLockValue(lock)) {
                continue;
            }
            String proc =
                    procCol >= 0 && procCol < row.size() && row.get(procCol) != null
                            ? row.get(procCol).strip()
                            : "";
            String mach =
                    machCol >= 0 && machCol < row.size() && row.get(machCol) != null
                            ? row.get(machCol).strip()
                            : "";
            labels.add("・" + proc + " × " + mach);
        }
        return labels;
    }

    private void applyCombinationRowVisibility() {
        if (comboView == null || comboView.getGrid() == null) {
            return;
        }
        BitSet extra =
                MasterDispatchSheetEditRules.combinationHiddenGridRows(
                        document.sheet("teamCombinations").rows(),
                        comboView.getGrid().getRowCount(),
                        SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex(),
                        equipmentFocusKeys);
        SpreadsheetMultiColumnFilterCoordinator.setAdditionalHiddenRows(comboView, extra);
    }

    private static void applyEquipmentColumnVisibility(
            SpreadsheetView view,
            MasterDispatchSheetEditRules.SheetKind kind,
            List<List<String>> rows,
            int leadingCols,
            Set<String> focusKeys) {
        if (view == null || view.getGrid() == null) {
            return;
        }
        int colCount = view.getGrid().getColumnCount();
        List<String> titles = MasterDispatchSheetEditRules.columnTitles(kind, rows, colCount);
        boolean[] vis = MasterDispatchSheetEditRules.visibilityMask(titles, leadingCols, focusKeys);
        ColumnVisibilitySupport.applyColumnVisibilityToSpreadsheetWhenReady(
                view,
                () -> titles,
                () -> vis,
                () -> {
                    List<Double> widths =
                            MasterDispatchSheetEditRules.preferredColumnWidths(
                                    rows, colCount, titles);
                    SpreadsheetTabularSupport.applyColumnWidths(view, widths, 120);
                    SpreadsheetTabularSupport.applyFixedLeadingColumns(view, leadingCols);
                });
    }

    private void attachGrid(
            SpreadsheetView view,
            MasterDispatchSheetEditRules.SheetKind kind,
            List<List<String>> rows,
            int frozenDataHeaderRows,
            int leadingCols) {
        attachGrid(view, kind, rows, frozenDataHeaderRows, leadingCols, List.of());
    }

    private void attachGrid(
            SpreadsheetView view,
            MasterDispatchSheetEditRules.SheetKind kind,
            List<List<String>> rows,
            int frozenDataHeaderRows,
            int leadingCols,
            List<List<String>> skillsRows) {
        List<List<String>> src = rows;
        if (kind == MasterDispatchSheetEditRules.SheetKind.COMBINATIONS) {
            src = MasterDispatchSheetEditRules.ensureCombinationMetaColumns(rows);
        }
        attachPrebuiltGrid(
                view,
                kind,
                MasterDispatchSheetGridSupport.buildEditable(kind, src, skillsRows),
                src,
                frozenDataHeaderRows,
                leadingCols);
    }

    private void attachPrebuiltGrid(
            SpreadsheetView view,
            MasterDispatchSheetEditRules.SheetKind kind,
            GridBase gridBase,
            List<List<String>> rows,
            int frozenDataHeaderRows,
            int leadingCols) {
        Grid previous = view.getGrid();
        if (previous != null) {
            previous.removeEventHandler(GridChange.GRID_CHANGE_EVENT, gridDirtyHandler);
        }
        view.setGrid(gridBase);
        Grid grid = view.getGrid();
        if (grid != null) {
            grid.addEventHandler(GridChange.GRID_CHANGE_EVENT, gridDirtyHandler);
        }
        int colCount = view.getGrid() != null ? view.getGrid().getColumnCount() : 1;
        List<String> titles = MasterDispatchSheetEditRules.columnTitles(kind, rows, colCount);
        List<Double> widths =
                MasterDispatchSheetEditRules.preferredColumnWidths(rows, colCount, titles);
        Platform.runLater(
                () -> applyMasterSheetChrome(view, widths, frozenDataHeaderRows, leadingCols));
    }

    private static void applyMasterSheetChrome(
            SpreadsheetView view,
            List<Double> widths,
            int frozenDataHeaderRows,
            int leadingCols) {
        SpreadsheetTabularSupport.applyColumnWidths(view, widths, 120);
        SpreadsheetTabularSupport.applyFixedLeadingColumns(view, leadingCols);
        SpreadsheetTabularSupport.applyColumnFiltersWithDialog(view);
        SpreadsheetMultiColumnFilterCoordinator.recomputeHiddenRows(view);
        SpreadsheetTabularSupport.pinSpreadsheetFilterRow(view);
        if (frozenDataHeaderRows > 0) {
            SpreadsheetTabularSupport.pinSpreadsheetRows(view, 1, frozenDataHeaderRows);
        }
        int lastTitle =
                SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex()
                        + Math.max(0, frozenDataHeaderRows)
                        - 1;
        SpreadsheetMultiColumnFilterCoordinator.setAlwaysVisibleThroughGridRow(view, lastTitle);
        SpreadsheetMultiColumnFilterCoordinator.recomputeHiddenRows(view);
        SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicyAfterSkinSettles(view);
        Platform.runLater(() -> SpreadsheetTabularSupport.applyColumnWidths(view, widths, 120));
    }

    private static String statusText(
            FactorySite site, Path json, Path source, MasterDispatchSheetsSeeder.Result result) {
        String factory = site != null ? site.displayLabelJa() : "";
        String outcome =
                switch (result.outcome()) {
                    case LOADED_EXISTING -> "既存 JSON を読み込みました。";
                    case IMPORTED -> "共有フォルダの master から吸い出して JSON を作成しました。";
                    case EMPTY_MISSING_SOURCE -> "吸い出し元の master が見つかりません。空の表を表示しています。";
                };
        return factory
                + "  JSON: "
                + json
                + "\n吸い出し元: "
                + source
                + "\n"
                + outcome;
    }
}
