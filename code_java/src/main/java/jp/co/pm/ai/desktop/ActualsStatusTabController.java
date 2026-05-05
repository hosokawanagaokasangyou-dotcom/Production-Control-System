package jp.co.pm.ai.desktop;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
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

    private static final String HINT_TEXT =
            "* データ行: 最列目を見出し除く累計"
                    + "（max "
                    + "500000"
                    + "行で扫描終了。実データ本体は"
                    + " Excel シート、本画面は JSON 形式の"
                    + " 取得メタのみ表示。";

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
                        new ColDef("種別", "label", 100),
                        new ColDef("解決根拠", "resolution", 160),
                        new ColDef("パス", "path", 220),
                        new ColDef("ファイル", "fileOk", 56),
                        new ColDef("シート", "sheetOk", 56),
                        new ColDef("データ行*", "dataRows", 88),
                        new ColDef("サイズ", "sizeBytes", 90),
                        new ColDef("更新(UTC)", "mtime", 140),
                        new ColDef("備考", "extra", 120));
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
        statusLabel.setText("取得中...");
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
                                        applyJson(cap.stdout(), rows, footLong, shell);
                                    });
                        });
    }

    private static void applyJson(String stdout, ObservableList<Row> rows, Text footLong, MainShellController shell) {
        rows.clear();
        String trimmed = stdout != null ? stdout.trim() : "";
        if (trimmed.isEmpty()) {
            footLong.setText("");
            shell.appendLog("[actuals-status] empty stdout");
            return;
        }
        try {
            JsonNode root = parseActualsPayloadRoot(trimmed);
            if (root.has("note")) {
                footLong.setText(root.get("note").asText(""));
            } else {
                footLong.setText("");
            }
            JsonNode entries = root.get("entries");
            if (entries == null || !entries.isArray()) {
                shell.appendLog("[actuals-status] no entries in JSON");
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
        } catch (Exception ex) {
            footLong.setText("");
            shell.appendLog("[actuals-status] JSON parse error: " + ex.getMessage());
        }
    }

    /**
     * Child stderr is merged into stdout; probe script prints one JSON line at the end. Parse
     * object-shaped lines from bottom to top, then fall back to the full buffer.
     */
    private static JsonNode parseActualsPayloadRoot(String trimmed) throws JsonProcessingException {
        String[] lines = trimmed.split("\\R", -1);
        JsonProcessingException lastLineFailure = null;
        for (int i = lines.length - 1; i >= 0; i--) {
            String ln = lines[i].trim();
            if (ln.isEmpty() || !ln.startsWith("{")) {
                continue;
            }
            try {
                return JSON.readTree(ln);
            } catch (JsonProcessingException e) {
                lastLineFailure = e;
            }
        }
        try {
            return JSON.readTree(trimmed);
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
