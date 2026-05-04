package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.Tab;
import javafx.scene.control.TabPane;
import javafx.scene.control.TextField;
import javafx.scene.control.Tooltip;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.StackPane;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.ui.GanttScheduleStyle;
import jp.co.pm.ai.desktop.ui.GanttSheetKind;
import jp.co.pm.ai.desktop.ui.SpreadsheetTabularSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetThemeBridge;

/**
 * {@code production_plan_multi_day*.json} \u3068 {@code member_schedule*.json} \u3092\u5165\u308c\u5b50
 * \u30bf\u30d6\uff08\u30c7\u30fc\u30bf\u30bb\u30c3\u30c8 \u2192 \u5404\u30b7\u30fc\u30c8 \u2192 \u8868/\u30ac\u30f3\u30c8\uff09\u3067\u8868\u793a\u3059\u308b\u3002
 */
public final class PlanResultViewerTabController {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final String HINT =
            "\u30c7\u30fc\u30bf\u30bb\u30c3\u30c8\u3092\u9078\u629e\u3057\u3001\u518d\u8aad\u307f\u3067\u5404\u30b7\u30fc\u30c8\u306e\u4e00\u89a7\u8868\u3068"
                    + "\u30ac\u30f3\u30c8\u98a8\u30bf\u30a4\u30e0\u30e9\u30a4\u30f3\uff08\u7a2e\u5225\u8272\uff09\u3092\u8868\u793a\u3057\u307e\u3059\u3002"
                    + " \u6700\u65b0JSON\u691c\u7d22\u306f PM_AI_OUTPUT_DIR \u4e0b\u306e\u6210\u679c\u7269\u30d5\u30a9\u30eb\u30c0\u304b\u3089\u6700\u65b0"
                    + " \u30da\u30a2\u3092\u63a2\u3057\u307e\u3059\u3002";

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
    private BorderPane contentPane;

    private MainShellController shell;

    private Stage ownerStage;

    /** Active spreadsheet views for \u5217\u30d5\u30a3\u30eb\u30bf\u89e3\u9664 */
    private final List<SpreadsheetView> registeredSpreadsheets = new ArrayList<>();

    @FXML
    private void initialize() {
        hintLabel.setText(HINT);
        contentPane.setCenter(emptyPlaceholder("\u518d\u8aad\u307f\u3067JSON\u3092\u8aad\u307f\u8fbc\u307f\u307e\u3059\u3002"));
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        this.ownerStage = shell.getPrimaryStage();
        Platform.runLater(this::reloadFromFields);
    }

    /**
     * \u5b9f\u884c\u30bf\u30d6\u306e\u6700\u65b0 xlsx \u30d1\u30b9\u3068\u540c\u3058\u30b9\u30c6\u30e0\u306e .json \u304c\u3042\u308c\u3070\u30d5\u30a3\u30fc\u30eb\u30c9\u306b\u53cd\u6620\uff08\u6bb5\u968e2\u5b8c\u4e86\u5f8c\u306a\u3069\uff09\u3002
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
            Path plan = newestMatching(dir, "production_plan_multi_day_*.json");
            Path mem = newestMatching(dir, "member_schedule_*.json");
            if (plan != null) {
                planJsonField.setText(plan.toString());
            }
            if (mem != null) {
                memberJsonField.setText(mem.toString());
            }
            if (plan == null && mem == null) {
                statusLabel.setText(
                        "\u3053\u306e\u30d5\u30a9\u30eb\u30c0\u306b JSON \u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093: "
                                + dir);
                return;
            }
        } catch (Exception ex) {
            statusLabel.setText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
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
                                "\u30d5\u30a1\u30a4\u30eb\u304c\u6307\u5b9a\u3055\u308c\u3066\u3044\u306a\u3044\u304b\u3001\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002"));
                statusLabel.setText("\u8aad\u307f\u8fbc\u307f\u5bfe\u8c61\u306a\u3057");
                return;
            }

            TabPane outer = new TabPane();
            outer.setTabClosingPolicy(TabPane.TabClosingPolicy.UNAVAILABLE);

            Tab tPlan = new Tab("\u751f\u7523\u8a08\u753b (multi_day)");
            TabPane planInner =
                    buildDatasetTabs(
                            planSheets,
                            planPath != null ? planPath.getFileName().toString() : "");
            tPlan.setContent(planInner);

            Tab tMem = new Tab("\u30e1\u30f3\u30d0\u30fc\u52e4\u52d9");
            TabPane memInner =
                    buildDatasetTabs(
                            memberSheets,
                            memberPath != null ? memberPath.getFileName().toString() : "");
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
                            + " \u8aad\u307f\u8fbc\u307f");
        } catch (Exception ex) {
            contentPane.setCenter(emptyPlaceholder("Error"));
            statusLabel.setText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
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

    private TabPane buildDatasetTabs(Map<String, SheetModel> sheets, String fileLabel) {
        TabPane inner = new TabPane();
        inner.setTabClosingPolicy(TabPane.TabClosingPolicy.UNAVAILABLE);
        if (sheets.isEmpty()) {
            inner.getTabs()
                    .add(
                            new Tab(
                                    "(\u7a7a)",
                                    emptyPlaceholder(
                                            "\u30c7\u30fc\u30bf\u306a\u3057 "
                                                    + fileLabel)));
            return inner;
        }
        for (Map.Entry<String, SheetModel> e : sheets.entrySet()) {
            String sheetName = e.getKey();
            SheetModel model = e.getValue();
            Tab st = new Tab(truncateTabTitle(sheetName));
            st.setTooltip(new Tooltip(sheetName + " \u2014 " + fileLabel));

            TabPane modeTabs = new TabPane();
            modeTabs.setTabClosingPolicy(TabPane.TabClosingPolicy.UNAVAILABLE);

            StackPane tableHost = new StackPane(new Label("\u8aad\u307f\u8fbc\u307f\u4e2d..."));
            StackPane ganttHost = new StackPane(new Label("\u8aad\u307f\u8fbc\u307f\u4e2d..."));

            Tab tTable = new Tab("\u4e00\u89a7\uff08\u8868\uff09", tableHost);
            Tab tGantt = new Tab("\u30ac\u30f3\u30c8", ganttHost);
            GanttSheetKind ganttKind =
                    GanttScheduleStyle.resolveKind(sheetName, model.columns());
            modeTabs.getTabs().addAll(tTable, tGantt);

            final boolean[] built = new boolean[2];
            Runnable loadTable =
                    () -> {
                        if (built[0]) {
                            return;
                        }
                        built[0] = true;
                        SpreadsheetView sv = new SpreadsheetView();
                        SpreadsheetThemeBridge.install(sv);
                        sv.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
                        ObservableList<ObservableList<String>> rows = model.copyRows();
                        GridBase grid =
                                SpreadsheetTabularSupport.buildReadOnlyPlainGrid(
                                        model.columns(), rows);
                        sv.setGrid(grid);
                        SpreadsheetTabularSupport.applyColumnFilters(sv);
                        SpreadsheetTabularSupport.applyFixedLeadingColumnsLater(sv, 1);
                        SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicy(sv);
                        Platform.runLater(
                                () -> {
                                    SpreadsheetTabularSupport.applyColumnWidths(
                                            sv, List.of(), 104);
                                    tableHost.getChildren().setAll(sv);
                                    StackPane.setAlignment(sv, Pos.CENTER_LEFT);
                                });
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
                        sv.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
                        ObservableList<ObservableList<String>> rows = model.copyRows();
                        GridBase grid =
                                SpreadsheetTabularSupport.buildReadOnlyGanttGrid(
                                        model.columns(), rows, ganttKind);
                        if (ganttKind == GanttSheetKind.EQUIPMENT_TIMELINE) {
                            sv.getStyleClass().add("pm-gantt-equipment-xlsx");
                        }
                        sv.setGrid(grid);
                        SpreadsheetTabularSupport.applyColumnFilters(sv);
                        SpreadsheetTabularSupport.applyFixedLeadingColumnsLater(sv, 1);
                        SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicy(sv);
                        Platform.runLater(
                                () -> {
                                    SpreadsheetTabularSupport.applyColumnWidths(
                                            sv, List.of(), 96);
                                    ganttHost.getChildren().setAll(sv);
                                    StackPane.setAlignment(sv, Pos.CENTER_LEFT);
                                });
                        registeredSpreadsheets.add(sv);
                    };
            st.setUserData(new Runnable[] {loadTable, loadGantt});

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

            st.setContent(modeTabs);
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

    private static void kickVisibleSheetTab(Tab sheetTab) {
        if (sheetTab == null || !(sheetTab.getContent() instanceof TabPane modeTabs)) {
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
        return s.length() <= max ? s : s.substring(0, max - 1) + "\u2026";
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

    private static Path newestMatching(Path dir, String glob) throws IOException {
        if (!Files.isDirectory(dir)) {
            return null;
        }
        Path best = null;
        long bestTime = Long.MIN_VALUE;
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(dir, glob)) {
            for (Path p : stream) {
                if (!Files.isRegularFile(p)) {
                    continue;
                }
                long t = Files.getLastModifiedTime(p).toMillis();
                if (t >= bestTime) {
                    bestTime = t;
                    best = p;
                }
            }
        }
        return best;
    }
}
