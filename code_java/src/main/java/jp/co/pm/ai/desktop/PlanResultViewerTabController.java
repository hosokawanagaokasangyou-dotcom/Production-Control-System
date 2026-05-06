package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.geometry.Side;
import javafx.scene.control.Accordion;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.RadioButton;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.Slider;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.Tab;
import javafx.scene.control.TabPane;
import javafx.scene.control.TextField;
import javafx.scene.control.TitledPane;
import javafx.scene.control.Tooltip;
import javafx.scene.Node;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.scene.text.Font;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.Stage2OutputNaming;
import jp.co.pm.ai.desktop.ui.GanttScheduleStyle;
import jp.co.pm.ai.desktop.ui.GanttSheetKind;
import jp.co.pm.ai.desktop.ui.SliderCommittedChangeSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnReorderDialog;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnSettingsStrip;
import jp.co.pm.ai.desktop.ui.SpreadsheetTabularSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetThemeBridge;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;

/**
 * {@code 計画*.json} と {@code 人員*.json}（旧英語名も解決可）を入れ子
 * タブ（データセット → 各シート → 表/ガント）で表示する。
 */
public final class PlanResultViewerTabController {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final String HINT =
            "データセットを選択し、再読みで各シートの一覧表と"
                    + "ガント風タイムライン（種別色）を表示します。"
                    + " 最新JSON検索は PM_AI_OUTPUT_DIR 下の成果物フォルダから最新"
                    + " ペアを探します。";

    @FXML
    private Button reloadButton;

    @FXML
    private Button syncLatestButton;

    @FXML
    private TextField planJsonField;

    @FXML
    private TextField memberJsonField;

    @FXML
    private Button browsePlanButton;

    @FXML
    private Button browseMemberButton;

    @FXML
    private Label statusLabel;

    @FXML
    private Label hintLabel;

    @FXML
    private ComboBox<String> planResultFontFamilyCombo;

    @FXML
    private Spinner<Integer> planResultFontSizeSpinner;

    @FXML
    private Slider planResultRowHeightSlider;

    @FXML
    private Label planResultRowHeightPctLabel;

    @FXML
    private RadioButton planResultCellWrapRadio;

    @FXML
    private RadioButton planResultCellClipRadio;

    @FXML
    private Accordion sourceAccordion;

    @FXML
    private TitledPane sourceTitledPane;

    @FXML
    private BorderPane contentPane;

    private MainShellController shell;

    private Stage ownerStage;

    /** Active spreadsheet views for 列フィルタ解除 */
    private final List<SpreadsheetView> registeredSpreadsheets = new ArrayList<>();

    private final AtomicReference<TableColumnOrderPersistence.PlanResultViewerUiPrefs> planResultUiPrefs =
            new AtomicReference<>(TableColumnOrderPersistence.loadPlanResultViewerUiPrefs());

    @FXML
    private void initialize() {
        hintLabel.setText(HINT);
        contentPane.setCenter(emptyPlaceholder("再読みでJSONを読み込みます。"));
        if (sourceAccordion != null && sourceTitledPane != null) {
            sourceAccordion.setExpandedPane(sourceTitledPane);
            sourceTitledPane.setExpanded(false);
        }
        initPlanResultUiControls();
    }

    private void initPlanResultUiControls() {
        if (planResultFontFamilyCombo == null || planResultFontSizeSpinner == null) {
            return;
        }
        TableColumnOrderPersistence.PlanResultViewerUiPrefs fp = planResultUiPrefs.get();
        ObservableList<String> families = FXCollections.observableArrayList(Font.getFamilies());
        FXCollections.sort(families);
        planResultFontFamilyCombo.setItems(families);
        String fam = fp.family() != null ? fp.family().strip() : "";
        if (!fam.isEmpty() && planResultFontFamilyCombo.getItems().contains(fam)) {
            planResultFontFamilyCombo.getSelectionModel().select(fam);
        } else {
            planResultFontFamilyCombo.getSelectionModel().select(Font.getDefault().getFamily());
        }
        int sz = (int) Math.round(fp.size());
        sz = Math.max(8, Math.min(48, sz <= 0 ? 12 : sz));
        planResultFontSizeSpinner.setValueFactory(new SpinnerValueFactory.IntegerSpinnerValueFactory(8, 48, sz));

        if (planResultRowHeightSlider != null) {
            double rh = fp.rowHeightPercent();
            if (Double.isNaN(rh) || rh < 50) {
                rh = 100.0;
            }
            rh =
                    Math.min(SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MAX, rh);
            planResultRowHeightSlider.setMin(SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MIN);
            planResultRowHeightSlider.setMax(SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MAX);
            planResultRowHeightSlider.setValue(rh);
            planResultRowHeightSlider.setMajorTickUnit(250);
            planResultRowHeightSlider.setMinorTickCount(4);
            planResultRowHeightSlider.setShowTickMarks(true);
            if (planResultRowHeightPctLabel != null) {
                planResultRowHeightPctLabel.setText(String.format("%.0f%%", rh));
            }
        }
        if (planResultCellWrapRadio != null && planResultCellClipRadio != null) {
            if (fp.cellWrapText()) {
                planResultCellWrapRadio.setSelected(true);
            } else {
                planResultCellClipRadio.setSelected(true);
            }
        }

        Runnable saveAndApply =
                () -> {
                    String selFam =
                            planResultFontFamilyCombo.getSelectionModel().getSelectedItem();
                    String effFam = selFam != null ? selFam.strip() : "";
                    Integer spv = planResultFontSizeSpinner.getValue();
                    double fs = spv != null ? spv.doubleValue() : 12.0;
                    double rowPct =
                            planResultRowHeightSlider != null
                                    ? planResultRowHeightSlider.getValue()
                                    : 100.0;
                    boolean wrap =
                            planResultCellWrapRadio != null && planResultCellWrapRadio.isSelected();
                    TableColumnOrderPersistence.PlanResultViewerUiPrefs next =
                            new TableColumnOrderPersistence.PlanResultViewerUiPrefs(
                                    effFam, fs, rowPct, wrap);
                    planResultUiPrefs.set(next);
                    TableColumnOrderPersistence.savePlanResultViewerUiPrefs(next);
                    if (planResultRowHeightPctLabel != null && planResultRowHeightSlider != null) {
                        planResultRowHeightPctLabel.setText(
                                String.format("%.0f%%", planResultRowHeightSlider.getValue()));
                    }
                    applyPlanResultPresentationToAllRegisteredSpreadsheets();
                };
        planResultFontFamilyCombo
                .getSelectionModel()
                .selectedItemProperty()
                .addListener((o, a, b) -> saveAndApply.run());
        planResultFontSizeSpinner
                .valueProperty()
                .addListener((o, a, b) -> saveAndApply.run());
        if (planResultRowHeightSlider != null) {
            SliderCommittedChangeSupport.install(
                    planResultRowHeightSlider,
                    () -> {
                        if (planResultRowHeightPctLabel != null) {
                            planResultRowHeightPctLabel.setText(
                                    String.format(
                                            "%.0f%%",
                                            planResultRowHeightSlider.getValue()));
                        }
                    },
                    saveAndApply::run);
        }
        if (planResultCellWrapRadio != null) {
            planResultCellWrapRadio.selectedProperty().addListener((o, a, b) -> saveAndApply.run());
        }
        if (planResultCellClipRadio != null) {
            planResultCellClipRadio.selectedProperty().addListener((o, a, b) -> saveAndApply.run());
        }
    }

    private void applyPlanResultFontStyle(SpreadsheetView sv) {
        if (sv == null) {
            return;
        }
        TableColumnOrderPersistence.PlanResultViewerUiPrefs p = planResultUiPrefs.get();
        double fsize = p.size() >= 6 ? p.size() : 12.0;
        fsize = Math.min(96, fsize);
        String fam = p.family() != null ? p.family().strip() : "";
        String esc = fam.replace("\\", "\\\\").replace("\"", "\\\"");
        String style =
                fam.isEmpty()
                        ? ("-fx-font-size: " + (int) Math.round(fsize) + "px;")
                        : ("-fx-font-family: \""
                                + esc
                                + "\"; -fx-font-size: "
                                + (int) Math.round(fsize)
                                + "px;");
        sv.setStyle(style);
    }

    private void applyPlanResultGridPresentation(SpreadsheetView sv) {
        if (sv == null) {
            return;
        }
        if (!(sv.getGrid() instanceof GridBase gb)) {
            return;
        }
        TableColumnOrderPersistence.PlanResultViewerUiPrefs u = planResultUiPrefs.get();
        SpreadsheetTabularSupport.applyPlanResultGridPresentation(
                gb, u.cellWrapText(), u.rowHeightPercent());
        SpreadsheetTabularSupport.refreshSpreadsheetAfterRowPresentationChange(sv);
    }

    /** フォントスタイルと行高・折り返しを現在の UI 設定で適用する。 */
    private void applyPlanResultPresentation(SpreadsheetView sv) {
        applyPlanResultFontStyle(sv);
        applyPlanResultGridPresentation(sv);
    }

    private void applyPlanResultPresentationToAllRegisteredSpreadsheets() {
        for (SpreadsheetView v : registeredSpreadsheets) {
            applyPlanResultPresentation(v);
        }
    }

    /** 読み込み成功後はデータ領域を広げるため折りたたむ */
    private void collapseSourcePaneAfterLoad() {
        if (sourceAccordion != null) {
            sourceAccordion.setExpandedPane(null);
        }
        if (sourceTitledPane != null) {
            sourceTitledPane.setExpanded(false);
        }
    }

    /** エラー時など、ファイル指定を見せる */
    private void expandSourcePaneForAttention() {
        if (sourceAccordion != null && sourceTitledPane != null) {
            sourceAccordion.setExpandedPane(sourceTitledPane);
            sourceTitledPane.setExpanded(true);
        }
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        this.ownerStage = shell.getPrimaryStage();
        Platform.runLater(this::reloadFromFields);
    }

    /**
     * 実行タブの最新 xlsx パスと同じステムの .json があればフィールドに反映（段階2完了後など）。
     */
    void tryAutofillJsonFromStage2Xlsx(String productionPlanPath, String memberSchedulePath) {
        if (planJsonField == null || memberJsonField == null) {
            return;
        }
        String p = productionPlanPath != null ? productionPlanPath.strip() : "";
        String m = memberSchedulePath != null ? memberSchedulePath.strip() : "";
        boolean touched = false;
        if (!p.isEmpty()) {
            Path json = siblingJson(Path.of(p));
            if (json != null && Files.isRegularFile(json)) {
                planJsonField.setText(json.toString());
                touched = true;
            }
        }
        if (!m.isEmpty()) {
            Path json = siblingJson(Path.of(m));
            if (json != null && Files.isRegularFile(json)) {
                memberJsonField.setText(json.toString());
                touched = true;
            }
        }
        if (touched) {
            Platform.runLater(this::reloadFromFields);
        }
    }

    @FXML
    private void onReloadButtonAction() {
        reloadFromFields();
    }

    @FXML
    private void onSyncLatestButtonAction() {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        Path dir = AppPaths.defaultPlanningOutputDir(ui);
        try {
            Path plan = Stage2OutputNaming.newestPrimaryPlanJson(dir);
            Path mem = Stage2OutputNaming.newestPrimaryMemberJson(dir);
            if (plan != null) {
                planJsonField.setText(plan.toString());
            }
            if (mem != null) {
                memberJsonField.setText(mem.toString());
            }
            if (plan == null && mem == null) {
                statusLabel.setText(
                        "このフォルダに JSON が見つかりません: "
                                + dir);
                expandSourcePaneForAttention();
                return;
            }
        } catch (Exception ex) {
            statusLabel.setText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
            expandSourcePaneForAttention();
            return;
        }
        reloadFromFields();
    }

    @FXML
    private void onBrowsePlanJsonAction() {
        chooseJson(planJsonField);
    }

    @FXML
    private void onBrowseMemberJsonAction() {
        chooseJson(memberJsonField);
    }

    private void chooseJson(TextField target) {
        FileChooser ch = new FileChooser();
        ch.setTitle("JSON");
        ch.getExtensionFilters().add(new FileChooser.ExtensionFilter("JSON", "*.json"));
        ch.getExtensionFilters().add(new FileChooser.ExtensionFilter("All", "*.*"));
        if (shell != null) {
            try {
                Map<String, String> ui = shell.snapshotUiEnv();
                Path dir = AppPaths.defaultPlanningOutputDir(ui);
                if (Files.isDirectory(dir)) {
                    ch.setInitialDirectory(dir.toFile());
                }
            } catch (Exception ignored) {
                // ignore
            }
        }
        java.io.File picked = ch.showOpenDialog(ownerStage);
        if (picked != null) {
            target.setText(picked.getAbsolutePath());
            reloadFromFields();
        }
    }

    void clearColumnFiltersAndSort() {
        for (SpreadsheetView v : registeredSpreadsheets) {
            SpreadsheetTabularSupport.clearAllFiltersAndSort(v);
        }
    }

    @FXML
    private void onClearColumnFiltersAction() {
        clearColumnFiltersAndSort();
    }

    private void reloadFromFields() {
        if (contentPane == null) {
            return;
        }
        registeredSpreadsheets.clear();
        reloadButton.setDisable(true);
        syncLatestButton.setDisable(true);
        try {
            String ps = planJsonField != null ? planJsonField.getText().strip() : "";
            String ms = memberJsonField != null ? memberJsonField.getText().strip() : "";
            Path planPath = ps.isEmpty() ? null : Path.of(ps);
            Path memberPath = ms.isEmpty() ? null : Path.of(ms);

            Map<String, SheetModel> planSheets =
                    planPath != null && Files.isRegularFile(planPath)
                            ? parseWorkbookSheets(planPath)
                            : Map.of();
            Map<String, SheetModel> memberSheets =
                    memberPath != null && Files.isRegularFile(memberPath)
                            ? parseWorkbookSheets(memberPath)
                            : Map.of();

            if (planSheets.isEmpty() && memberSheets.isEmpty()) {
                contentPane.setCenter(
                        emptyPlaceholder(
                                "ファイルが指定されていないか、見つかりません。"));
                statusLabel.setText("読み込み対象なし");
                expandSourcePaneForAttention();
                return;
            }

            TabPane outer = new TabPane();
            outer.setSide(Side.TOP);
            outer.getStyleClass().add("pm-plan-result-tabpane-side");
            outer.setTabClosingPolicy(TabPane.TabClosingPolicy.UNAVAILABLE);

            Tab tPlan = new Tab("生産計画 (multi_day)");
            TabPane planInner =
                    buildDatasetTabs(
                            planSheets,
                            planPath != null ? planPath.getFileName().toString() : "",
                            "plan");
            tPlan.setContent(planInner);

            Tab tMem = new Tab("メンバー勤務");
            TabPane memInner =
                    buildDatasetTabs(
                            memberSheets,
                            memberPath != null ? memberPath.getFileName().toString() : "",
                            "member");
            tMem.setContent(memInner);

            outer.getTabs().add(tPlan);
            outer.getTabs().add(tMem);
            wireDatasetTabActivation(outer);
            contentPane.setCenter(outer);
            Platform.runLater(() -> kickVisibleSheetLoaders(outer.getSelectionModel().getSelectedItem()));

            statusLabel.setText(
                    "plan_sheets="
                            + planSheets.size()
                            + ", member_sheets="
                            + memberSheets.size()
                            + " 読み込み");
            collapseSourcePaneAfterLoad();
        } catch (Exception ex) {
            contentPane.setCenter(emptyPlaceholder("Error"));
            statusLabel.setText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
            expandSourcePaneForAttention();
            if (shell != null) {
                shell.appendLog("[plan-result-viewer] " + ex.getMessage());
            }
        } finally {
            reloadButton.setDisable(false);
            syncLatestButton.setDisable(false);
        }
    }

    private StackPane emptyPlaceholder(String msg) {
        StackPane p = new StackPane(new Label(msg));
        StackPane.setAlignment(p.getChildren().get(0), Pos.CENTER);
        return p;
    }

    private TabPane buildDatasetTabs(
            Map<String, SheetModel> sheets, String fileLabel, String datasetTag) {
        TabPane inner = new TabPane();
        inner.setSide(Side.TOP);
        inner.getStyleClass().add("pm-plan-result-tabpane-side");
        inner.setTabClosingPolicy(TabPane.TabClosingPolicy.UNAVAILABLE);
        if (sheets.isEmpty()) {
            inner.getTabs()
                    .add(
                            new Tab(
                                    "(空)",
                                    emptyPlaceholder(
                                            "データなし "
                                                    + fileLabel)));
            return inner;
        }
        for (Map.Entry<String, SheetModel> e : sheets.entrySet()) {
            String sheetName = e.getKey();
            SheetModel model = e.getValue();
            Tab st = new Tab(truncateTabTitle(sheetName));
            st.setTooltip(new Tooltip(sheetName + " — " + fileLabel));

            SheetGridState gridState = SheetGridState.fromSheetModel(model, datasetTag, sheetName);

            TabPane modeTabs = new TabPane();
            modeTabs.setTabClosingPolicy(TabPane.TabClosingPolicy.UNAVAILABLE);

            StackPane tableHost = new StackPane(new Label("読み込み中..."));
            StackPane ganttHost = new StackPane(new Label("読み込み中..."));

            GanttSheetKind ganttKind =
                    GanttScheduleStyle.resolveKind(sheetName, gridState.headersRef);

            Tab tTable = new Tab("一覧（表）", tableHost);
            Tab tGantt = new Tab("ガント", ganttHost);
            modeTabs.getTabs().addAll(tTable, tGantt);

            AtomicReference<SpreadsheetView> tableSvRef = new AtomicReference<>();
            AtomicReference<SpreadsheetView> ganttSvRef = new AtomicReference<>();
            AtomicBoolean tableWatcherInstalled = new AtomicBoolean();
            AtomicBoolean ganttWatcherInstalled = new AtomicBoolean();
            final boolean[] built = new boolean[2];

            Runnable rebuildTable =
                    () -> {
                        SpreadsheetView sv = tableSvRef.get();
                        if (sv == null) {
                            return;
                        }
                        gridState.suppressPersistence.set(true);
                        try {
                            GridBase grid =
                                    SpreadsheetTabularSupport.buildReadOnlyPlainGrid(
                                            gridState.headersRef, gridState.rows);
                            sv.setGrid(grid);
                            Platform.runLater(
                                    () ->
                                            finishPlanResultSpreadsheet(
                                                    sv, gridState, tableWatcherInstalled));
                        } finally {
                            gridState.suppressPersistence.set(false);
                        }
                    };

            Runnable rebuildGantt =
                    () -> {
                        SpreadsheetView sv = ganttSvRef.get();
                        if (sv == null) {
                            return;
                        }
                        gridState.suppressPersistence.set(true);
                        try {
                            GridBase grid =
                                    SpreadsheetTabularSupport.buildReadOnlyGanttGrid(
                                            gridState.headersRef, gridState.rows, ganttKind);
                            sv.setGrid(grid);
                            Platform.runLater(
                                    () -> {
                                        if (ganttKind == GanttSheetKind.EQUIPMENT_TIMELINE
                                                && !sv.getStyleClass()
                                                        .contains("pm-gantt-equipment-xlsx")) {
                                            sv.getStyleClass().add("pm-gantt-equipment-xlsx");
                                        }
                                        finishPlanResultSpreadsheet(
                                                sv, gridState, ganttWatcherInstalled);
                                    });
                        } finally {
                            gridState.suppressPersistence.set(false);
                        }
                    };

            Runnable resetWidths =
                    () -> {
                        SpreadsheetView v = tableSvRef.get();
                        if (v == null) {
                            v = ganttSvRef.get();
                        }
                        if (v == null) {
                            return;
                        }
                        double def = 112;
                        for (var c : v.getColumns()) {
                            c.setPrefWidth(def);
                        }
                        TableColumnOrderPersistence.saveLayoutForScope(
                                gridState.scopeKey,
                                TableColumnOrderPersistence.snapshotSpreadsheet(
                                        v, gridState.headersRef));
                    };

            Runnable onReorder =
                    () -> {
                        if (gridState.headersRef.isEmpty()) {
                            if (shell != null) {
                                shell.appendLog("[plan-result-viewer] 列がありません（先にシートを開く）");
                            }
                            return;
                        }
                        SpreadsheetColumnReorderDialog.show(
                                        ownerStage, new ArrayList<>(gridState.headersRef))
                                .ifPresent(
                                        perm -> {
                                            List<String> oldHeaders =
                                                    new ArrayList<>(gridState.headersRef);
                                            List<String> titleOrder =
                                                    perm.stream().map(oldHeaders::get).toList();
                                            TableColumnOrderPersistence.applyLogicalColumnOrder(
                                                    gridState.headersRef,
                                                    gridState.rows,
                                                    titleOrder);
                                            List<Double> widths =
                                                    TableColumnOrderPersistence.resolveWidthsForHeaders(
                                                            gridState.headersRef,
                                                            gridState.persistedLayout.get(),
                                                            112);
                                            List<TableColumnOrderPersistence.ColumnSpec> newLay =
                                                    new ArrayList<>();
                                            for (int i = 0; i < gridState.headersRef.size(); i++) {
                                                newLay.add(
                                                        new TableColumnOrderPersistence.ColumnSpec(
                                                                gridState.headersRef.get(i),
                                                                widths.get(i)));
                                            }
                                            gridState.persistedLayout.set(newLay);
                                            TableColumnOrderPersistence.saveLayoutForScope(
                                                    gridState.scopeKey, newLay);
                                            if (built[0]) {
                                                rebuildTable.run();
                                            }
                                            if (built[1]) {
                                                rebuildGantt.run();
                                            }
                                        });
                    };

            HBox strip =
                    SpreadsheetColumnSettingsStrip.createForScope(
                            resetWidths,
                            gridState.scopeKey,
                            gridState.headerColumnCount,
                            hv -> {
                                if (built[0]) {
                                    rebuildTable.run();
                                }
                                if (built[1]) {
                                    rebuildGantt.run();
                                }
                            },
                            onReorder);

            Runnable loadTable =
                    () -> {
                        if (built[0]) {
                            return;
                        }
                        built[0] = true;
                        SpreadsheetView sv = new SpreadsheetView();
                        SpreadsheetThemeBridge.install(sv);
                        applyPlanResultPresentation(sv);
                        sv.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
                        tableSvRef.set(sv);
                        gridState.suppressPersistence.set(true);
                        try {
                            GridBase grid =
                                    SpreadsheetTabularSupport.buildReadOnlyPlainGrid(
                                            gridState.headersRef, gridState.rows);
                            sv.setGrid(grid);
                            Platform.runLater(
                                    () -> {
                                        finishPlanResultSpreadsheet(
                                                sv, gridState, tableWatcherInstalled);
                                        tableHost.getChildren().setAll(sv);
                                        StackPane.setAlignment(sv, Pos.CENTER_LEFT);
                                    });
                        } finally {
                            gridState.suppressPersistence.set(false);
                        }
                        registeredSpreadsheets.add(sv);
                    };
            Runnable loadGantt =
                    () -> {
                        if (built[1]) {
                            return;
                        }
                        built[1] = true;
                        SpreadsheetView sv = new SpreadsheetView();
                        SpreadsheetThemeBridge.install(sv);
                        applyPlanResultPresentation(sv);
                        sv.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
                        ganttSvRef.set(sv);
                        gridState.suppressPersistence.set(true);
                        try {
                            GridBase grid =
                                    SpreadsheetTabularSupport.buildReadOnlyGanttGrid(
                                            gridState.headersRef, gridState.rows, ganttKind);
                            sv.setGrid(grid);
                            Platform.runLater(
                                    () -> {
                                        if (ganttKind == GanttSheetKind.EQUIPMENT_TIMELINE) {
                                            sv.getStyleClass().add("pm-gantt-equipment-xlsx");
                                        }
                                        finishPlanResultSpreadsheet(
                                                sv, gridState, ganttWatcherInstalled);
                                        ganttHost.getChildren().setAll(sv);
                                        StackPane.setAlignment(sv, Pos.CENTER_LEFT);
                                    });
                        } finally {
                            gridState.suppressPersistence.set(false);
                        }
                        registeredSpreadsheets.add(sv);
                    };
            st.setUserData(new Runnable[] {loadTable, loadGantt});

            Accordion columnSettingsAccordion = new Accordion();
            TitledPane columnSettingsPane = new TitledPane("列設定", strip);
            columnSettingsPane.setExpanded(false);
            columnSettingsAccordion.getPanes().add(columnSettingsPane);
            columnSettingsAccordion.setMaxWidth(Double.MAX_VALUE);
            columnSettingsAccordion.getStyleClass().add("pm-plan-result-sheet-column-accordion");

            VBox sheetVBox = new VBox(4, columnSettingsAccordion, modeTabs);
            VBox.setVgrow(columnSettingsAccordion, Priority.NEVER);
            VBox.setVgrow(modeTabs, Priority.ALWAYS);

            modeTabs
                    .getSelectionModel()
                    .selectedIndexProperty()
                    .addListener(
                            (obs, o, n) -> {
                                if (n == null) {
                                    return;
                                }
                                int idx = n.intValue();
                                if (!(st.getUserData() instanceof Runnable[] loaders)) {
                                    return;
                                }
                                if (idx < 0 || idx >= loaders.length) {
                                    return;
                                }
                                loaders[idx].run();
                            });

            st.selectedProperty()
                    .addListener(
                            (obs, o, now) -> {
                                if (Boolean.TRUE.equals(now)) {
                                    int m = modeTabs.getSelectionModel().getSelectedIndex();
                                    if (st.getUserData() instanceof Runnable[] loaders
                                            && m >= 0
                                            && m < loaders.length) {
                                        loaders[m].run();
                                    }
                                }
                            });

            st.setContent(sheetVBox);
            inner.getTabs().add(st);
        }
        inner.getSelectionModel()
                .selectedItemProperty()
                .addListener(
                        (obs, o, n) -> {
                            if (n != null) {
                                kickVisibleSheetTab(n);
                            }
                        });
        return inner;
    }

    private void finishPlanResultSpreadsheet(
            SpreadsheetView sv,
            SheetGridState gridState,
            AtomicBoolean layoutWatcherInstalled) {
        List<Double> widths =
                TableColumnOrderPersistence.resolveWidthsForHeaders(
                        gridState.headersRef, gridState.persistedLayout.get(), 112);
        SpreadsheetTabularSupport.applyColumnWidths(sv, widths, 112);
        SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicy(sv);
        SpreadsheetTabularSupport.applyColumnFiltersWithDialog(sv);
        SpreadsheetTabularSupport.applyFixedLeadingColumnsLater(
                sv, gridState.headerColumnCount.get());
        applyPlanResultPresentation(sv);
        if (layoutWatcherInstalled.compareAndSet(false, true)) {
            TableColumnOrderPersistence.installSpreadsheetColumnLayoutWatcherForScope(
                    sv,
                    gridState.scopeKey,
                    gridState.suppressPersistence::get,
                    () -> new ArrayList<>(gridState.headersRef));
        }
    }

    private static final class SheetGridState {
        final ArrayList<String> headersRef = new ArrayList<>();
        final ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        final AtomicReference<List<TableColumnOrderPersistence.ColumnSpec>> persistedLayout =
                new AtomicReference<>(List.of());
        final AtomicInteger headerColumnCount = new AtomicInteger(0);
        final AtomicBoolean suppressPersistence = new AtomicBoolean(false);
        final String scopeKey;

        private SheetGridState(String scopeKey) {
            this.scopeKey = scopeKey;
        }

        static SheetGridState fromSheetModel(
                SheetModel model, String datasetTag, String sheetName) {
            String scope =
                    TableColumnOrderPersistence.planResultViewerSheetScopeKey(datasetTag, sheetName);
            SheetGridState st = new SheetGridState(scope);
            st.headersRef.addAll(model.columns());
            for (Map<String, String> map : model.rowMaps()) {
                ObservableList<String> line = FXCollections.observableArrayList();
                for (String h : st.headersRef) {
                    line.add(map.getOrDefault(h, ""));
                }
                st.rows.add(line);
            }
            List<TableColumnOrderPersistence.ColumnSpec> lay =
                    TableColumnOrderPersistence.loadLayoutForScope(scope);
            st.persistedLayout.set(lay);
            TableColumnOrderPersistence.applyLogicalColumnOrder(
                    st.headersRef,
                    st.rows,
                    lay.stream().map(TableColumnOrderPersistence.ColumnSpec::title).toList());
            st.headerColumnCount.set(TableColumnOrderPersistence.loadHeaderColumnCountForScope(scope));
            return st;
        }
    }

    private static void wireDatasetTabActivation(TabPane outer) {
        outer.getSelectionModel()
                .selectedItemProperty()
                .addListener(
                        (obs, o, n) -> kickVisibleSheetLoaders(n));
    }

    private static void kickVisibleSheetLoaders(Tab datasetTab) {
        if (datasetTab == null || !(datasetTab.getContent() instanceof TabPane inner)) {
            return;
        }
        Tab sheet = inner.getSelectionModel().getSelectedItem();
        if (sheet != null) {
            kickVisibleSheetTab(sheet);
        }
    }

    private static TabPane findModeTabPaneInSheetTab(Tab sheetTab) {
        if (sheetTab == null) {
            return null;
        }
        Node content = sheetTab.getContent();
        if (content instanceof TabPane tp) {
            return tp;
        }
        if (content instanceof VBox v) {
            for (Node n : v.getChildren()) {
                if (n instanceof TabPane inner) {
                    return inner;
                }
            }
        }
        return null;
    }

    private static void kickVisibleSheetTab(Tab sheetTab) {
        TabPane modeTabs = findModeTabPaneInSheetTab(sheetTab);
        if (modeTabs == null) {
            return;
        }
        Object ud = sheetTab.getUserData();
        if (!(ud instanceof Runnable[] loaders) || loaders.length < 1) {
            return;
        }
        int mi = modeTabs.getSelectionModel().getSelectedIndex();
        if (mi < 0 || mi >= loaders.length) {
            return;
        }
        loaders[mi].run();
    }

    private static String truncateTabTitle(String s) {
        if (s == null) {
            return "";
        }
        int max = 18;
        return s.length() <= max ? s : s.substring(0, max - 1) + "…";
    }

    private static Map<String, SheetModel> parseWorkbookSheets(Path path) throws IOException {
        JsonNode root = JSON.readTree(Files.readString(path, StandardCharsets.UTF_8));
        JsonNode sheetsNode = root.get("sheets");
        if (sheetsNode == null || !sheetsNode.isObject()) {
            throw new IOException("JSON: sheets object missing");
        }
        Map<String, SheetModel> out = new LinkedHashMap<>();
        Iterator<Map.Entry<String, JsonNode>> it = sheetsNode.fields();
        while (it.hasNext()) {
            Map.Entry<String, JsonNode> en = it.next();
            SheetModel m = parseSheetModel(en.getValue());
            if (m != null) {
                out.put(en.getKey(), m);
            }
        }
        return out;
    }

    private static SheetModel parseSheetModel(JsonNode sheetNode) {
        if (sheetNode == null || !sheetNode.isObject()) {
            return null;
        }
        JsonNode columnsNode = sheetNode.get("columns");
        JsonNode rowsNode = sheetNode.get("rows");
        if (columnsNode == null
                || !columnsNode.isArray()
                || rowsNode == null
                || !rowsNode.isArray()) {
            return null;
        }
        List<String> columns = new ArrayList<>();
        for (JsonNode c : columnsNode) {
            columns.add(c != null && c.isTextual() ? c.asText("") : "");
        }
        List<Map<String, String>> rowMaps = new ArrayList<>();
        for (JsonNode r : rowsNode) {
            if (r == null || !r.isObject()) {
                continue;
            }
            Map<String, String> row = new LinkedHashMap<>();
            for (String h : columns) {
                row.put(h, formatCell(r.get(h)));
            }
            rowMaps.add(row);
        }
        return new SheetModel(columns, rowMaps);
    }

    private static String formatCell(JsonNode n) {
        if (n == null || n.isNull()) {
            return "";
        }
        if (n.isBoolean()) {
            return n.asBoolean() ? "true" : "false";
        }
        if (n.isInt() || n.isLong()) {
            return Long.toString(n.longValue());
        }
        if (n.isDouble() || n.isFloat() || n.isBigDecimal()) {
            double d = n.asDouble();
            if (Double.isFinite(d) && d == Math.rint(d) && Math.abs(d) < 1e15) {
                return Long.toString((long) d);
            }
            return n.asText("");
        }
        if (n.isTextual()) {
            String t = n.asText("");
            if (t.length() >= 19 && t.charAt(10) == 'T' && t.charAt(4) == '-') {
                return t.substring(0, 10);
            }
            return t;
        }
        return n.asText("");
    }

    private record SheetModel(List<String> columns, List<Map<String, String>> rowMaps) {
        ObservableList<ObservableList<String>> copyRows() {
            ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
            for (Map<String, String> map : rowMaps) {
                ObservableList<String> line = FXCollections.observableArrayList();
                for (String h : columns) {
                    line.add(map.getOrDefault(h, ""));
                }
                rows.add(line);
            }
            return rows;
        }
    }

    private static Path siblingJson(Path workbookPath) {
        if (workbookPath == null) {
            return null;
        }
        Path fn = workbookPath.getFileName();
        if (fn == null) {
            return null;
        }
        String name = fn.toString();
        String stem;
        if (name.endsWith(".xlsx")) {
            stem = name.substring(0, name.length() - 5);
        } else if (name.endsWith(".xlsm")) {
            stem = name.substring(0, name.length() - 5);
        } else {
            return null;
        }
        return workbookPath.resolveSibling(stem + ".json");
    }
}
