package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.function.Consumer;
import java.util.function.Supplier;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.scene.Parent;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.scene.text.Text;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.bridge.PythonProcessRunner.CapturedResult;
import jp.co.pm.ai.desktop.bridge.PythonProcessRunner.RunRequest;

/**
 * Resolved paths and sheet scan metadata for machining actuals / actual-detail (Excel;
 * JSON status line from {@code pm_ai_actuals_status.py}).
 */
public final class ActualsDataStatusPane {

    private static final ObjectMapper JSON = new ObjectMapper();

    private ActualsDataStatusPane() {}

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

    public static Parent create(Supplier<RunRequest> requestFactory, Consumer<String> appendLog) {
        ObservableList<Row> rows = FXCollections.observableArrayList();
        TableView<Row> table = new TableView<>(rows);
        table.setColumnResizePolicy(TableView.UNCONSTRAINED_RESIZE_POLICY);

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
            col.setMinWidth(d.width());
            col.setPrefWidth(d.width());
            columns.add(col);
        }
        table.getColumns().setAll(columns);
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

        Platform.runLater(
                () ->
                        ColumnVisibilitySupport.applyColumnVisibilityToTableView(
                                table,
                                TableColumnOrderPersistence.loadColumnVisibility(
                                        TableColumnOrderPersistence.TableId.ACTUALS_STATUS,
                                        table.getColumns().size())));

        Runnable resetActualsColumns =
                () -> {
                    for (int i = 0; i < columns.size(); i++) {
                        double w = defs.get(i).width();
                        columns.get(i).setMinWidth(w);
                        columns.get(i).setPrefWidth(w);
                    }
                };
        HBox actualsColStrip =
                TableViewColumnSettingsStrip.create(
                        table,
                        resetActualsColumns,
                        false,
                        TableColumnOrderPersistence.TableId.ACTUALS_STATUS,
                        null,
                        () ->
                                ColumnVisibilitySupport.openTableViewColumnVisibilityDialog(
                                        table.getScene() != null ? table.getScene().getWindow() : null,
                                        TableColumnOrderPersistence.TableId.ACTUALS_STATUS,
                                        table));

        Text footLong = new Text();
        footLong.setWrappingWidth(880);

        Label status = new Label("");

        Button refresh = new Button("\u66f4\u65b0");
        refresh.setOnAction(
                ev -> {
                    refresh.setDisable(true);
                    status.setText("\u53d6\u5f97\u4e2d...");
                    RunRequest req = requestFactory.get();
                    PythonProcessRunner.runCaptureAsync(req)
                            .whenComplete(
                                    (cap, err) -> {
                                        Platform.runLater(
                                                () -> {
                                                    refresh.setDisable(false);
                                                    if (err != null) {
                                                        status.setText(
                                                                "error: "
                                                                        + err.getMessage());
                                                        appendLog.accept(
                                                                "[actuals-status] " + err.getMessage());
                                                        return;
                                                    }
                                                    if (cap == null) {
                                                        status.setText("no result");
                                                        return;
                                                    }
                                                    status.setText(
                                                            "exit="
                                                                    + cap.exitCode());
                                                    applyJson(cap.stdout(), rows, footLong, appendLog);
                                                });
                                    });
                });

        HBox bar = new HBox(8, refresh, status);
        bar.setPadding(new Insets(0, 0, 8, 0));

        Label hint =
                new Label(
                        "* \u30c7\u30fc\u30bf\u884c: \u6700\u5217\u76ee\u3092\u898b\u51fa\u3057\u9664\u304f\u7d2f\u8a08"
                                + "\uff08max "
                                + "500000"
                                + "\u884c\u3067\u626b\u63cf\u7d42\u4e86\u3002\u5b9f\u30c7\u30fc\u30bf\u672c\u4f53\u306f"
                                + " Excel \u30b7\u30fc\u30c8\u3001\u672c\u753b\u9762\u306f JSON \u5f62\u5f0f\u306e"
                                + " \u53d6\u5f97\u30e1\u30bf\u306e\u307f\u8868\u793a\u3002");
        hint.setWrapText(true);

        VBox top = new VBox(8, bar, hint, actualsColStrip);
        VBox bottom = new VBox(4, new Label("\u8aac\u660e"), footLong);
        VBox root = new VBox(8, top, table, bottom);
        root.setFillWidth(true);
        VBox.setVgrow(table, Priority.ALWAYS);
        VBox.setMargin(bottom, new Insets(8, 0, 0, 0));
        root.setPadding(new Insets(12));
        return root;
    }

    private static void applyJson(
            String stdout,
            ObservableList<Row> rows,
            Text footLong,
            Consumer<String> appendLog) {
        rows.clear();
        String trimmed = stdout != null ? stdout.trim() : "";
        if (trimmed.isEmpty()) {
            footLong.setText("");
            appendLog.accept("[actuals-status] empty stdout");
            return;
        }
        try {
            JsonNode root = JSON.readTree(trimmed);
            if (root.has("note")) {
                footLong.setText(root.get("note").asText(""));
            } else {
                footLong.setText("");
            }
            JsonNode entries = root.get("entries");
            if (entries == null || !entries.isArray()) {
                appendLog.accept("[actuals-status] no entries in JSON");
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
            appendLog.accept("[actuals-status] JSON parse error: " + ex.getMessage());
        }
    }

    private static String textOr(JsonNode n, String field) {
        JsonNode x = n.get(field);
        return x == null || x.isNull() ? "" : x.asText("");
    }
}
