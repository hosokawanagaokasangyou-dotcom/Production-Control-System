package jp.co.pm.ai.desktop;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.Supplier;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.scene.text.Text;

import org.controlsfx.control.table.TableFilter;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.bridge.PythonProcessRunner.RunRequest;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;
import jp.co.pm.ai.desktop.ui.TableHeaderColumnStyle;
import jp.co.pm.ai.desktop.ui.TableViewColumnSettingsStrip;

/** Actuals DATA status probe UI; layout {@code ActualsStatusTab.fxml}. */
public final class ActualsStatusTabController {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final Path AGENT_DEBUG_LOG =
            Path.of("/mnt/c/\u5de5\u7a0b\u7ba1\u7406AI\u30d7\u30ed\u30b8\u30a7\u30af\u30c8_JAVA/.cursor/debug-e25361.log");

    private static void agentLog(String hypothesisId, String message, Map<String, Object> data) {
        // #region agent log
        try {
            Map<String, Object> root = new HashMap<>();
            root.put("sessionId", "e25361");
            root.put("runId", "actuals-tab");
            root.put("hypothesisId", hypothesisId);
            root.put("location", "ActualsStatusTabController");
            root.put("message", message);
            root.put("data", data);
            root.put("timestamp", System.currentTimeMillis());
            String line = JSON.writeValueAsString(root) + "\n";
            Files.writeString(
                    AGENT_DEBUG_LOG,
                    line,
                    StandardCharsets.UTF_8,
                    StandardOpenOption.CREATE,
                    StandardOpenOption.APPEND);
        } catch (Throwable ignored) {
            // debug ingest only
        }
        // #endregion
    }

    private static final String HINT_TEXT =
            "* \u30c7\u30fc\u30bf\u884c: \u6700\u5217\u76ee\u3092\u898b\u51fa\u3057\u9664\u304f\u7d2f\u8a08"
                    + "\uff08max "
                    + "500000"
                    + "\u884c\u3067\u626b\u63cf\u7d42\u4e86\u3002\u5b9f\u30c7\u30fc\u30bf\u672c\u4f53\u306f"
                    + " Excel \u30b7\u30fc\u30c8\u3001\u672c\u753b\u9762\u306f JSON \u5f62\u5f0f\u306e"
                    + " \u53d6\u5f97\u30e1\u30bf\u306e\u307f\u8868\u793a\u3002";

    public static final class Row {
        private String label;
        private String resolution;
        private String path;
        private String fileOk;
        private String sheetOk;
        private String dataRows;
        private String sizeBytes;
        private String mtime;
        private String extra;

        public String getLabel() {
            return label;
        }

        public void setLabel(String label) {
            this.label = label;
        }

        public String getResolution() {
            return resolution;
        }

        public void setResolution(String resolution) {
            this.resolution = resolution;
        }

        public String getPath() {
            return path;
        }

        public void setPath(String path) {
            this.path = path;
        }

        public String getFileOk() {
            return fileOk;
        }

        public void setFileOk(String fileOk) {
            this.fileOk = fileOk;
        }

        public String getSheetOk() {
            return sheetOk;
        }

        public void setSheetOk(String sheetOk) {
            this.sheetOk = sheetOk;
        }

        public String getDataRows() {
            return dataRows;
        }

        public void setDataRows(String dataRows) {
            this.dataRows = dataRows;
        }

        public String getSizeBytes() {
            return sizeBytes;
        }

        public void setSizeBytes(String sizeBytes) {
            this.sizeBytes = sizeBytes;
        }

        public String getMtime() {
            return mtime;
        }

        public void setMtime(String mtime) {
            this.mtime = mtime;
        }

        public String getExtra() {
            return extra;
        }

        public void setExtra(String extra) {
            this.extra = extra;
        }
    }

    private record ColDef(String title, String prop, double width) {}

    @FXML
    private Button refreshButton;

    @FXML
    private Label statusLabel;

    @FXML
    private Label hintLabel;

    @FXML
    private HBox columnStripHost;

    @FXML
    private TableView<Row> table;

    @FXML
    private Text footLong;

    private ObservableList<Row> rows;

    private MainShellController shell;

    private Supplier<RunRequest> actualsStatusRequestFactory;

    private final AtomicInteger headerColumnCount = new AtomicInteger(0);

    private TableFilter<Row> tableFilter;

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
        this.actualsStatusRequestFactory = shell::buildActualsStatusRequest;
        List<ColDef> defs =
                Arrays.asList(
                        new ColDef("\u7a2e\u5225", "label", 100),
                        new ColDef("\u89e3\u6c7a\u6839\u62e0", "resolution", 160),
                        new ColDef("\u30d1\u30b9", "path", 220),
                        new ColDef("\u30d5\u30a1\u30a4\u30eb", "fileOk", 56),
                        new ColDef("\u30b7\u30fc\u30c8", "sheetOk", 56),
                        new ColDef("\u30c7\u30fc\u30bf\u884c*", "dataRows", 88),
                        new ColDef("\u30b5\u30a4\u30ba", "sizeBytes", 90),
                        new ColDef("\u66f4\u65b0(UTC)", "mtime", 140),
                        new ColDef("\u5099\u8003", "extra", 120));
        List<TableColumn<Row, String>> columns = new ArrayList<>();
        for (ColDef d : defs) {
            TableColumn<Row, String> col = new TableColumn<>(d.title());
            col.setCellValueFactory(new PropertyValueFactory<>(d.prop()));
            col.setCellFactory(
                    tc ->
                            new TableCell<Row, String>() {
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
            col.setMinWidth(d.width());
            col.setPrefWidth(d.width());
            col.setReorderable(true);
            columns.add(col);
        }
        table.getColumns().setAll(columns);
        tableFilter = TableFilter.forTableView(table).apply();
        List<TableColumnOrderPersistence.ColumnSpec> actualsLayout =
                TableColumnOrderPersistence.loadLayout(TableColumnOrderPersistence.TableId.ACTUALS_STATUS);
        if (!actualsLayout.isEmpty()) {
            TableColumnOrderPersistence.applyOrderToTableColumns(
                    table,
                    actualsLayout.stream()
                            .map(TableColumnOrderPersistence.ColumnSpec::title)
                            .toList());
            TableColumnOrderPersistence.applyWidthsToTableColumns(table, actualsLayout, 112);
        }
        TableColumnOrderPersistence.installColumnLayoutWatcher(
                table, TableColumnOrderPersistence.TableId.ACTUALS_STATUS, () -> false);

        Runnable resetActualsColumns =
                () -> {
                    for (int i = 0; i < columns.size(); i++) {
                        double w = defs.get(i).width();
                        columns.get(i).setMinWidth(w);
                        columns.get(i).setPrefWidth(w);
                    }
                };
        columnStripHost.getChildren()
                .setAll(
                        TableViewColumnSettingsStrip.create(
                                table,
                                resetActualsColumns,
                                false,
                                TableColumnOrderPersistence.TableId.ACTUALS_STATUS,
                                headerColumnCount));
    }

    @FXML
    private void onRefreshButtonAction() {
        refreshButton.setDisable(true);
        statusLabel.setText("\u53d6\u5f97\u4e2d...");
        RunRequest req = actualsStatusRequestFactory.get();
        PythonProcessRunner.runCaptureAsync(req)
                .whenComplete(
                        (cap, err) -> {
                            Platform.runLater(
                                    () -> {
                                        refreshButton.setDisable(false);
                                        if (err != null) {
                                            statusLabel.setText("error: " + err.getMessage());
                                            shell.appendLog("[actuals-status] " + err.getMessage());
                                            return;
                                        }
                                        if (cap == null) {
                                            statusLabel.setText("no result");
                                            return;
                                        }
                                        statusLabel.setText("exit=" + cap.exitCode());
                                        // #region agent log
                                        {
                                            String out = cap.stdout();
                                            Map<String, Object> d = new HashMap<>();
                                            d.put("exitCode", cap.exitCode());
                                            d.put("stdoutNull", out == null);
                                            d.put("stdoutLength", out != null ? out.length() : 0);
                                            d.put(
                                                    "stdoutFirst200",
                                                    out == null
                                                            ? ""
                                                            : out.substring(0, Math.min(200, out.length())));
                                            agentLog("A", "after python capture (actuals status)", d);
                                        }
                                        // #endregion
                                        applyJson(cap.stdout(), rows, footLong, shell);
                                    });
                        });
    }

    private static void applyJson(String stdout, ObservableList<Row> rows, Text footLong, MainShellController shell) {
        rows.clear();
        String trimmed = stdout != null ? stdout.trim() : "";
        // #region agent log
        {
            Map<String, Object> d = new HashMap<>();
            d.put("trimmedEmpty", trimmed.isEmpty());
            d.put("trimmedLength", trimmed.length());
            d.put("lineCount", trimmed.isEmpty() ? 0 : trimmed.split("\n", -1).length);
            d.put(
                    "trimmedFirst120",
                    trimmed.isEmpty() ? "" : trimmed.substring(0, Math.min(120, trimmed.length())));
            agentLog("B", "applyJson input", d);
        }
        // #endregion
        if (trimmed.isEmpty()) {
            footLong.setText("");
            shell.appendLog("[actuals-status] empty stdout");
            return;
        }
        try {
            ParseAttempt pa = parseActualsPayloadRoot(trimmed);
            JsonNode root = pa.root();
            // #region agent log
            {
                JsonNode ent = root.get("entries");
                Map<String, Object> d = new HashMap<>();
                List<String> keys = new ArrayList<>();
                root.fieldNames().forEachRemaining(keys::add);
                d.put("rootKeys", keys);
                d.put("entriesNull", ent == null || ent.isNull());
                d.put("entriesIsArray", ent != null && ent.isArray());
                d.put("entriesSize", ent != null && ent.isArray() ? ent.size() : -1);
                d.put("parseStrategy", pa.strategy());
                agentLog("C", "applyJson parsed root", d);
            }
            // #endregion
            if (root.has("note")) {
                footLong.setText(root.get("note").asText(""));
            } else {
                footLong.setText("");
            }
            JsonNode entries = root.get("entries");
            if (entries == null || !entries.isArray()) {
                shell.appendLog("[actuals-status] no entries in JSON");
                // #region agent log
                agentLog(
                        "C",
                        "applyJson missing or non-array entries",
                        Map.of("hadEntriesKey", root.has("entries")));
                // #endregion
                return;
            }
            List<Row> list = new ArrayList<>();
            for (JsonNode e : entries) {
                Row r = new Row();
                r.setLabel(textOr(e, "label"));
                r.setResolution(textOr(e, "resolution"));
                r.setPath(textOr(e, "resolved_path"));
                boolean fileExists = e.path("file_exists").asBoolean(false);
                r.setFileOk(fileExists ? "OK" : "-");
                boolean sheetFound = e.path("sheet_found").asBoolean(false);
                r.setSheetOk(sheetFound ? "OK" : "-");
                if (e.hasNonNull("data_rows")) {
                    String dr = e.get("data_rows").asText();
                    if (e.path("scan_truncated").asBoolean(false)) {
                        dr = dr + "+";
                    }
                    r.setDataRows(dr);
                } else {
                    r.setDataRows("-");
                }
                if (e.hasNonNull("size_bytes")) {
                    r.setSizeBytes(e.get("size_bytes").asText());
                } else {
                    r.setSizeBytes("-");
                }
                r.setMtime(textOr(e, "mtime_iso"));
                StringBuilder ex = new StringBuilder();
                if (e.hasNonNull("error")) {
                    ex.append(e.get("error").asText());
                }
                if (e.hasNonNull("sheet_name")) {
                    if (ex.length() > 0) {
                        ex.append(" ");
                    }
                    ex.append("sheet=").append(e.get("sheet_name").asText());
                }
                r.setExtra(ex.toString());
                list.add(r);
            }
            rows.setAll(list);
            // #region agent log
            agentLog("E", "applyJson rows set", Map.of("rowCount", list.size()));
            // #endregion
        } catch (Exception ex) {
            footLong.setText("");
            shell.appendLog("[actuals-status] JSON parse error: " + ex.getMessage());
            // #region agent log
            Map<String, Object> d = new HashMap<>();
            d.put("exception", ex.getClass().getName());
            d.put("message", ex.getMessage());
            d.put(
                    "trimmedFirst200",
                    trimmed.isEmpty() ? "" : trimmed.substring(0, Math.min(200, trimmed.length())));
            agentLog("A", "applyJson parse/IO exception", d);
            // #endregion
        }
    }

    /**
     * Child stderr is merged into stdout; probe script prints one JSON line at the end. Parse
     * object-shaped lines from bottom to top, then fall back to the full buffer.
     */
    private record ParseAttempt(JsonNode root, String strategy) {}

    private static ParseAttempt parseActualsPayloadRoot(String trimmed) throws JsonProcessingException {
        String[] lines = trimmed.split("\\R", -1);
        JsonProcessingException lastLineFailure = null;
        for (int i = lines.length - 1; i >= 0; i--) {
            String ln = lines[i].trim();
            if (ln.isEmpty() || !ln.startsWith("{")) {
                continue;
            }
            try {
                return new ParseAttempt(JSON.readTree(ln), "lastJsonLine index=" + i);
            } catch (JsonProcessingException e) {
                lastLineFailure = e;
            }
        }
        try {
            return new ParseAttempt(JSON.readTree(trimmed), "fullBuffer");
        } catch (JsonProcessingException e) {
            if (lastLineFailure != null) {
                throw lastLineFailure;
            }
            throw e;
        }
    }

    private static String textOr(JsonNode n, String field) {
        JsonNode x = n.get(field);
        return x == null || x.isNull() ? "" : x.asText("");
    }

    void clearColumnFiltersAndSort() {
        if (tableFilter != null) {
            tableFilter.resetAllFilters();
        }
        table.getSortOrder().clear();
    }
}
