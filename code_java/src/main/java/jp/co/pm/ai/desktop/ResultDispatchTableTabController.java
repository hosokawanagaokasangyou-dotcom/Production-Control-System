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

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import javafx.application.Platform;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.scene.text.Text;

import org.controlsfx.control.table.TableFilter;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;
import jp.co.pm.ai.desktop.ui.TableHeaderColumnStyle;
import jp.co.pm.ai.desktop.ui.TableViewColumnSettingsStrip;

/** Loads {@link AppPaths#RESULT_DISPATCH_TABLE_JSON_BASENAME} as a wide table. Layout {@code ResultDispatchTableTab.fxml}. */
public final class ResultDispatchTableTabController {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final String HINT_TEXT =
            "PM_AI_RESULT_DISPATCH_TABLE_DIR \u307e\u305f\u306f\u30c7\u30d5\u30a9\u30eb\u30c8\u306e code/"
                    + " \u914d\u4e0b\u306e JSON \u3092\u8868\u793a\u3057\u307e\u3059\u3002\u518d\u8aad\u307f\u3067"
                    + "\u6700\u65b0\u5316\u3057\u307e\u3059\u3002";

    private record ColSpec(String key, double width) {}

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
    private TableView<Map<String, String>> table;

    @FXML
    private Text metaText;

    private ObservableList<Map<String, String>> rows;

    private MainShellController shell;

    private TableFilter<Map<String, String>> tableFilter;

    private final AtomicInteger headerColumnCount = new AtomicInteger(0);

    private final AtomicBoolean suppressColumnPersistence = new AtomicBoolean(false);

    private List<ColSpec> columnSpecs = List.of();

    private boolean layoutWatcherInstalled;

    private boolean columnStripInstalled;

    @FXML
    private void initialize() {
        hintLabel.setText(HINT_TEXT);
        rows = FXCollections.observableArrayList();
        table.setItems(rows);
        table.setColumnResizePolicy(TableView.UNCONSTRAINED_RESIZE_POLICY);
        VBox.setVgrow(table, Priority.ALWAYS);
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        Platform.runLater(this::reloadFromDisk);
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
            applyEmptyTable();
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
                applyEmptyTable();
                refreshButton.setDisable(false);
                return;
            }
            List<String> headers = new ArrayList<>();
            for (JsonNode c : columnsNode) {
                if (c != null && c.isTextual()) {
                    headers.add(c.asText(""));
                }
            }
            List<Map<String, String>> data = new ArrayList<>();
            for (JsonNode r : rowsNode) {
                if (r == null || !r.isObject()) {
                    continue;
                }
                LinkedHashMap<String, String> row = new LinkedHashMap<>();
                for (String h : headers) {
                    row.put(h, formatCell(r.get(h)));
                }
                data.add(row);
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
                            + data.size();
            metaText.setText(meta);
            statusLabel.setText(data.size() + " \u884c");
            rebuildTable(headers, data);
        } catch (Exception ex) {
            statusLabel.setText("error");
            metaText.setText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
            shell.appendLog("[result-dispatch-json] " + ex.getMessage());
            applyEmptyTable();
        } finally {
            refreshButton.setDisable(false);
        }
    }

    private void applyEmptyTable() {
        suppressColumnPersistence.set(true);
        try {
            rows.clear();
            table.getColumns().clear();
            if (tableFilter != null) {
                tableFilter = null;
            }
            columnSpecs = List.of();
        } finally {
            suppressColumnPersistence.set(false);
        }
    }

    private void rebuildTable(List<String> headers, List<Map<String, String>> data) {
        suppressColumnPersistence.set(true);
        try {
            rows.clear();
            table.getColumns().clear();
            if (tableFilter != null) {
                tableFilter = null;
            }
            if (headers.isEmpty()) {
                columnSpecs = List.of();
                return;
            }
            double baseW = 112.0;
            List<TableColumn<Map<String, String>, String>> cols = new ArrayList<>();
            List<ColSpec> specs = new ArrayList<>();
            for (String h : headers) {
                final String key = h;
                double w = guessColumnWidth(key, baseW);
                specs.add(new ColSpec(key, w));
                TableColumn<Map<String, String>, String> col = new TableColumn<>(key);
                col.setCellValueFactory(
                        cd -> {
                            Map<String, String> row = cd.getValue();
                            String v = row != null ? row.getOrDefault(key, "") : "";
                            return new SimpleStringProperty(v != null ? v : "");
                        });
                col.setCellFactory(
                        tc ->
                                new TableCell<Map<String, String>, String>() {
                                    @Override
                                    protected void updateItem(String item, boolean empty) {
                                        super.updateItem(item, empty);
                                        if (empty || item == null) {
                                            setText(null);
                                        } else {
                                            setText(item);
                                        }
                                        TableHeaderColumnStyle.applyBodyCellTint(
                                                this, table, tc, headerColumnCount::get);
                                    }
                                });
                col.setMinWidth(w);
                col.setPrefWidth(w);
                col.setReorderable(true);
                cols.add(col);
            }
            columnSpecs = List.copyOf(specs);
            table.getColumns().addAll(cols);
            rows.setAll(data);
            tableFilter = TableFilter.forTableView(table).apply();
            List<TableColumnOrderPersistence.ColumnSpec> saved =
                    TableColumnOrderPersistence.loadLayout(TableColumnOrderPersistence.TableId.RESULT_DISPATCH_TABLE);
            if (!saved.isEmpty()) {
                TableColumnOrderPersistence.applyOrderToTableColumns(
                        table,
                        saved.stream().map(TableColumnOrderPersistence.ColumnSpec::title).toList());
                TableColumnOrderPersistence.applyWidthsToTableColumns(table, saved, baseW);
            }
            if (!columnStripInstalled) {
                Runnable resetWidths =
                        () -> {
                            for (int i = 0; i < columnSpecs.size() && i < table.getColumns().size(); i++) {
                                double w = columnSpecs.get(i).width();
                                TableColumn<Map<String, String>, ?> c = table.getColumns().get(i);
                                c.setMinWidth(w);
                                c.setPrefWidth(w);
                            }
                        };
                columnStripHost.getChildren()
                        .setAll(
                                TableViewColumnSettingsStrip.create(
                                        table,
                                        resetWidths,
                                        false,
                                        TableColumnOrderPersistence.TableId.RESULT_DISPATCH_TABLE,
                                        headerColumnCount));
                columnStripInstalled = true;
            }
            if (!layoutWatcherInstalled) {
                TableColumnOrderPersistence.installColumnLayoutWatcher(
                        table,
                        TableColumnOrderPersistence.TableId.RESULT_DISPATCH_TABLE,
                        suppressColumnPersistence::get);
                layoutWatcherInstalled = true;
            }
        } finally {
            suppressColumnPersistence.set(false);
        }
    }

    private static double guessColumnWidth(String header, double base) {
        if (header == null) {
            return base;
        }
        int len = header.length();
        if (len <= 4) {
            return Math.min(220, base + 16);
        }
        if (len >= 12) {
            return Math.min(280, base + 48);
        }
        return base;
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
        if (tableFilter != null) {
            tableFilter.resetAllFilters();
        }
        table.getSortOrder().clear();
    }
}
