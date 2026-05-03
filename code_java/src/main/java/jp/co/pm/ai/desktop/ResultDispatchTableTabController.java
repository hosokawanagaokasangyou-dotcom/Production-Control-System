package jp.co.pm.ai.desktop;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
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
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.SelectionMode;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.scene.text.Text;
import javafx.stage.Stage;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnReorderDialog;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnSettingsStrip;
import jp.co.pm.ai.desktop.ui.SpreadsheetTabularSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetThemeBridge;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;

/**
 * Loads {@link AppPaths#RESULT_DISPATCH_TABLE_JSON_BASENAME} into ControlsFX {@link SpreadsheetView}. Layout
 * {@code ResultDispatchTableTab.fxml}.
 */
public final class ResultDispatchTableTabController {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final String HINT_TEXT =
            "PM_AI_RESULT_DISPATCH_TABLE_DIR \u307e\u305f\u306f\u30c7\u30d5\u30a9\u30eb\u30c8\u306e code/"
                    + " \u914d\u4e0b\u306e JSON \u3092\u8868\u793a\u3057\u307e\u3059\u3002\u518d\u8aad\u307f\u3067"
                    + "\u6700\u65b0\u5316\u3057\u307e\u3059\u3002"
                    + " ControlsFX SpreadsheetView \uff08\u6bb5\u968e1\u6210\u5f62\u7d50\u679c\u3068\u540c\u3058"
                    + "\u5217\u30d5\u30a3\u30eb\u30bf\uff09\u3002";

    @FXML
    private Button refreshButton;

    @FXML
    private Label statusLabel;

    @FXML
    private Label pathLabel;

    @FXML
    private Label hintLabel;

    @FXML
    private HBox columnStripHost;

    @FXML
    private StackPane spreadsheetHost;

    @FXML
    private Text metaText;

    private MainShellController shell;

    private Stage ownerStage;

    private final SpreadsheetView spreadsheetView = new SpreadsheetView();

    private final List<String> headersRef = new ArrayList<>();

    private ObservableList<ObservableList<String>> rows;

    private final AtomicBoolean suppressColumnPersistence = new AtomicBoolean(false);

    private final AtomicReference<List<TableColumnOrderPersistence.ColumnSpec>> persistedLayout =
            new AtomicReference<>(List.of());

    private final AtomicInteger headerColumnCount = new AtomicInteger(0);

    @FXML
    private void initialize() {
        hintLabel.setText(HINT_TEXT);
        rows = FXCollections.observableArrayList();

        StackPane.setAlignment(spreadsheetView, Pos.CENTER_LEFT);
        spreadsheetHost.getChildren().add(spreadsheetView);
        VBox.setVgrow(spreadsheetHost, Priority.ALWAYS);

        spreadsheetView.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        SpreadsheetThemeBridge.install(spreadsheetView);

        columnStripHost
                .getChildren()
                .setAll(
                        SpreadsheetColumnSettingsStrip.create(
                                this::applyDynamicColumnWidths,
                                TableColumnOrderPersistence.TableId.RESULT_DISPATCH_TABLE,
                                headerColumnCount,
                                this::onLeadingColumnCountCommitted,
                                this::onReorderColumns));
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        ownerStage = shell.getPrimaryStage();

        TableColumnOrderPersistence.installSpreadsheetColumnLayoutWatcher(
                spreadsheetView,
                TableColumnOrderPersistence.TableId.RESULT_DISPATCH_TABLE,
                suppressColumnPersistence::get,
                () -> new ArrayList<>(headersRef));

        Platform.runLater(this::reloadFromDisk);
    }

    private void onLeadingColumnCountCommitted(int n) {
        headerColumnCount.set(n);
        rebuildSpreadsheet();
    }

    private void onReorderColumns() {
        if (headersRef.isEmpty()) {
            shell.appendLog(
                    "[result-dispatch-json] \u5217\u304c\u3042\u308a\u307e\u305b\u3093\uff08\u5148\u306b\u518d\u8aad\u307f\uff09");
            return;
        }
        SpreadsheetColumnReorderDialog.show(ownerStage, new ArrayList<>(headersRef))
                .ifPresent(
                        perm -> {
                            List<String> oldHeaders = new ArrayList<>(headersRef);
                            List<String> titleOrder = perm.stream().map(oldHeaders::get).toList();
                            List<TableColumnOrderPersistence.ColumnSpec> lay = persistedLayout.get();
                            TableColumnOrderPersistence.applyLogicalColumnOrder(
                                    headersRef, rows, titleOrder);
                            List<Double> widths =
                                    TableColumnOrderPersistence.resolveWidthsForHeaders(
                                            headersRef, lay, 112);
                            List<TableColumnOrderPersistence.ColumnSpec> newLay = new ArrayList<>();
                            for (int i = 0; i < headersRef.size(); i++) {
                                newLay.add(
                                        new TableColumnOrderPersistence.ColumnSpec(
                                                headersRef.get(i), widths.get(i)));
                            }
                            persistedLayout.set(newLay);
                            TableColumnOrderPersistence.saveLayout(
                                    TableColumnOrderPersistence.TableId.RESULT_DISPATCH_TABLE, newLay);
                            rebuildSpreadsheet();
                        });
    }

    private void applyDynamicColumnWidths() {
        double w = 112;
        for (var c : spreadsheetView.getColumns()) {
            c.setPrefWidth(w);
        }
    }

    private void rebuildSpreadsheet() {
        if (headersRef.isEmpty()) {
            spreadsheetView.setGrid(new GridBase(0, 0));
            return;
        }
        suppressColumnPersistence.set(true);
        try {
            final List<Double> widths =
                    TableColumnOrderPersistence.resolveWidthsForHeaders(
                            headersRef, persistedLayout.get(), 112);
            final double widthDefault = 112;

            GridBase grid = SpreadsheetTabularSupport.buildReadOnlyPlainGrid(headersRef, rows);
            spreadsheetView.setGrid(grid);

            Platform.runLater(
                    () -> {
                        SpreadsheetTabularSupport.applyColumnWidths(
                                spreadsheetView, widths, widthDefault);
                        SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicy(spreadsheetView);
                        SpreadsheetTabularSupport.applyFixedLeadingColumns(
                                spreadsheetView, headerColumnCount.get());
                        SpreadsheetTabularSupport.applyColumnFilters(spreadsheetView);
                    });
        } finally {
            suppressColumnPersistence.set(false);
        }
    }

    @FXML
    private void onRefreshButtonAction() {
        reloadFromDisk();
    }

    private void reloadFromDisk() {
        if (shell == null) {
            return;
        }
        refreshButton.setDisable(true);
        Map<String, String> ui = shell.snapshotUiEnv();
        Path path = AppPaths.resolveResultDispatchTableJsonPath(ui);
        pathLabel.setText(path.toString());
        if (!Files.isRegularFile(path)) {
            statusLabel.setText("\u30d5\u30a1\u30a4\u30eb\u306a\u3057");
            metaText.setText("");
            applyEmpty();
            refreshButton.setDisable(false);
            return;
        }
        try {
            String raw = Files.readString(path, StandardCharsets.UTF_8);
            JsonNode root = JSON.readTree(raw);
            String sheetName = textOr(root, "sheet_name");
            String excelTable = textOr(root, "excel_table_name");
            int formatVer = root.path("format_version").asInt(0);
            int rowCountMeta = root.path("row_count").asInt(-1);
            JsonNode columnsNode = root.get("columns");
            JsonNode rowsNode = root.get("rows");
            if (columnsNode == null || !columnsNode.isArray() || rowsNode == null || !rowsNode.isArray()) {
                statusLabel.setText("JSON \u69cb\u9020\u304c\u4e0d\u6b63");
                metaText.setText("");
                applyEmpty();
                refreshButton.setDisable(false);
                return;
            }
            List<String> headerOrder = new ArrayList<>();
            for (JsonNode c : columnsNode) {
                if (c != null && c.isTextual()) {
                    headerOrder.add(c.asText(""));
                }
            }
            List<Map<String, String>> rowMaps = new ArrayList<>();
            for (JsonNode r : rowsNode) {
                if (r == null || !r.isObject()) {
                    continue;
                }
                LinkedHashMap<String, String> row = new LinkedHashMap<>();
                for (String h : headerOrder) {
                    row.put(h, formatCell(r.get(h)));
                }
                rowMaps.add(row);
            }
            String meta =
                    "format_version="
                            + formatVer
                            + ", sheet_name="
                            + sheetName
                            + ", excel_table_name="
                            + excelTable
                            + ", row_count="
                            + rowCountMeta
                            + ", loaded_rows="
                            + rowMaps.size();
            metaText.setText(meta);
            statusLabel.setText(rowMaps.size() + " \u884c");

            headersRef.clear();
            headersRef.addAll(headerOrder);
            rows.clear();
            for (Map<String, String> map : rowMaps) {
                ObservableList<String> line = FXCollections.observableArrayList();
                for (String h : headersRef) {
                    line.add(map.getOrDefault(h, ""));
                }
                rows.add(line);
            }

            List<TableColumnOrderPersistence.ColumnSpec> lay =
                    TableColumnOrderPersistence.loadLayout(
                            TableColumnOrderPersistence.TableId.RESULT_DISPATCH_TABLE);
            persistedLayout.set(lay);
            TableColumnOrderPersistence.applyLogicalColumnOrder(
                    headersRef,
                    rows,
                    lay.stream().map(TableColumnOrderPersistence.ColumnSpec::title).toList());

            rebuildSpreadsheet();
        } catch (Exception ex) {
            statusLabel.setText("error");
            metaText.setText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
            shell.appendLog("[result-dispatch-json] " + ex.getMessage());
            applyEmpty();
        } finally {
            refreshButton.setDisable(false);
        }
    }

    private void applyEmpty() {
        headersRef.clear();
        rows.clear();
        persistedLayout.set(List.of());
        spreadsheetView.setGrid(new GridBase(0, 0));
    }

    private static String textOr(JsonNode n, String field) {
        JsonNode x = n.get(field);
        return x == null || x.isNull() ? "" : x.asText("");
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

    void clearColumnFiltersAndSort() {
        SpreadsheetTabularSupport.clearAllFiltersAndSort(spreadsheetView);
    }
}
