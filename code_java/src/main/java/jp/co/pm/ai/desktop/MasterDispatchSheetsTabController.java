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
import javafx.fxml.FXML;
import javafx.geometry.Insets;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Label;
import javafx.scene.control.TabPane;
import javafx.scene.layout.StackPane;
import javafx.stage.Window;

import org.controlsfx.control.spreadsheet.SpreadsheetView;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.io.MasterDispatchSheetsDocument;
import jp.co.pm.ai.desktop.io.MasterDispatchSheetsSaveWriter;
import jp.co.pm.ai.desktop.io.MasterDispatchSheetsSeeder;
import jp.co.pm.ai.desktop.io.MasterTeamCombinationTableReader;
import jp.co.pm.ai.desktop.ui.ColumnVisibilityDialog;
import jp.co.pm.ai.desktop.ui.ColumnVisibilitySupport;
import jp.co.pm.ai.desktop.ui.FourDigitConfirmationDialog;
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
    @FXML private TabPane innerTabPane;
    @FXML private StackPane skillsHost;
    @FXML private StackPane needHost;
    @FXML private StackPane speedHost;
    @FXML private StackPane comboHost;

    private final SpreadsheetView skillsView = new SpreadsheetView();
    private final SpreadsheetView needView = new SpreadsheetView();
    private final SpreadsheetView speedView = new SpreadsheetView();
    private final SpreadsheetView comboView = new SpreadsheetView();

    private MainShellController shell;
    private MasterDispatchSheetsDocument document = MasterDispatchSheetsDocument.empty("");
    private Path loadedJsonPath;
    /** 4タブ共通。工程+機械の正規化キー。空なら空き列以外をすべて表示。 */
    private final Set<String> equipmentFocusKeys = new LinkedHashSet<>();

    @FXML
    private void initialize() {
        install(skillsHost, skillsView, 1);
        install(needHost, needView, 3);
        install(speedHost, speedView, 3);
        install(comboHost, comboView, 4);
        if (innerTabPane != null) {
            innerTabPane
                    .getSelectionModel()
                    .selectedIndexProperty()
                    .addListener((obs, o, n) -> applyAllEquipmentVisibility());
        }
        applyDocument(document);
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
        SpreadsheetTabularSupport.installSpreadsheetChromeRelayoutDebouncerForHost(host, () -> leadingCols);
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
    }

    void onMainShellTabSelected() {
        reloadFromCurrentFactory(false);
    }

    void reloadFromCurrentFactory(boolean reimport) {
        equipmentFocusKeys.clear();
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        FactorySite site = AppPaths.currentDispatchFactorySite(ui);
        Path json = AppPaths.masterDispatchSheetsJsonPath(ui);
        Path source = AppPaths.masterDispatchSheetsSourceWorkbookPath(ui);
        try {
            MasterDispatchSheetsSeeder.Result result =
                    MasterDispatchSheetsSeeder.loadOrImport(json, source, site.name(), reimport);
            document = result.document();
            loadedJsonPath = json;
            applyDocument(document);
            statusLabel.setText(statusText(site, json, source, result));
        } catch (Exception e) {
            statusLabel.setText("読込に失敗しました: " + e.getMessage());
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
            alert.setTitle("MASTER 保存");
            alert.setHeaderText("入力内容を直してから保存してください。");
            alert.setContentText(String.join("\n", errors.subList(0, Math.min(8, errors.size()))));
            Window w = statusLabel != null && statusLabel.getScene() != null ? statusLabel.getScene().getWindow() : null;
            if (w != null) {
                alert.initOwner(w);
            }
            alert.showAndWait();
            statusLabel.setText("保存を中止しました。検証エラーがあります。");
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
                        + "保存前に JSON と master.xlsm の世代バックアップを取ります。",
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
        equipmentFocusKeys.clear();
        int added = applyEquipmentColumn(r.process(), r.machine());
        focusSkillsTab();
        if (added <= 0) {
            statusLabel.setText("既に同じ工程名+機械名の列があります: " + r.process() + " × " + r.machine());
            return;
        }
        statusLabel.setText(
                "列を追加しました: "
                        + r.process()
                        + " × "
                        + r.machine()
                        + "。追加した列だけ表示しています。OP/AS を設定し、保存してください。「すべて表示」で全列に戻せます。");
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
        int added = 0;
        for (PlanTasksMissingSkillsColumnPrompt.MissingPair pair : bundle.pairs()) {
            added += applyEquipmentColumn(pair.process(), pair.machine());
        }
        focusSkillsTab();
        statusLabel.setText(
                "未登録の工程+機械を "
                        + added
                        + " 列追加しました。追加した列だけ表示しています。OP/AS を設定し、保存してください。「すべて表示」で全列に戻せます。");
        return added;
    }

    int addMissingEquipmentColumns(List<PlanTasksMissingSkillsColumnPrompt.MissingPair> pairs) {
        if (pairs == null || pairs.isEmpty()) {
            return 0;
        }
        equipmentFocusKeys.clear();
        int added = 0;
        for (PlanTasksMissingSkillsColumnPrompt.MissingPair pair : pairs) {
            added += applyEquipmentColumn(pair.process(), pair.machine());
        }
        focusSkillsTab();
        statusLabel.setText(
                "未登録の工程+機械を "
                        + added
                        + " 列追加しました。追加した列だけ表示しています。OP/AS を設定し、保存してください。「すべて表示」で全列に戻せます。");
        return added;
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
            if (!key.isEmpty()) {
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
                4);
        applyAllEquipmentVisibility();
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
                view, () -> titles, () -> vis);
    }

    private static void attachGrid(
            SpreadsheetView view,
            MasterDispatchSheetEditRules.SheetKind kind,
            List<List<String>> rows,
            int frozenDataHeaderRows,
            int leadingCols) {
        view.setGrid(MasterDispatchSheetGridSupport.buildEditable(kind, rows));
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
