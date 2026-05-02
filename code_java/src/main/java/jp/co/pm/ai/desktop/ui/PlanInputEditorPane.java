package jp.co.pm.ai.desktop.ui;

import java.io.File;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;
import java.util.function.Supplier;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.scene.Parent;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.io.PlanInputTabularIo;

/**
 * Edits {@code \u914d\u53f0\u8a08\u753b_\u30bf\u30b9\u30af\u5165\u529b}-equivalent tabular data (CSV / xlsx).
 * Align with {@code PM_AI_PLAN_INPUT_PATH} and optional {@code TASK_PLAN_SHEET} (sheet name in Excel).
 */
public final class PlanInputEditorPane {

    public static final String ENV_PM_AI_PLAN_INPUT_PATH = "PM_AI_PLAN_INPUT_PATH";
    public static final String ENV_TASK_PLAN_SHEET = "TASK_PLAN_SHEET";

    /** Default sheet name (same as planning_core PLAN_INPUT_SHEET_NAME when TASK_PLAN_SHEET is empty). */
    public static final String DEFAULT_PLAN_INPUT_SHEET_NAME = "\u914d\u53f0\u8a08\u753b_\u30bf\u30b9\u30af\u5165\u529b";

    private PlanInputEditorPane() {}

    public static Parent create(
            Stage owner,
            Supplier<Map<String, String>> envSupplier,
            Consumer<String> log) {

        TextField pathField = new TextField();
        pathField.setPromptText("PM_AI_PLAN_INPUT_PATH ? .csv / .xlsx / .xlsm");

        TextField sheetField = new TextField(DEFAULT_PLAN_INPUT_SHEET_NAME);
        sheetField.setPromptText("Excel sheet name (TASK_PLAN_SHEET / TASK_PLAN_SHEET)");

        TableView<ObservableList<String>> table = new TableView<>();
        table.setEditable(true);
        table.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        table.setItems(rows);

        TextField colWidthField = new TextField("112");
        colWidthField.setMaxWidth(72);
        colWidthField.setPromptText("112");

        List<String> headersRef = new ArrayList<>();

        Runnable applyDynamicColumnWidths =
                () -> {
                    double w = 112;
                    try {
                        w = Math.max(40, Double.parseDouble(colWidthField.getText().trim()));
                    } catch (NumberFormatException ignored) {
                    }
                    for (TableColumn<ObservableList<String>, ?> c : table.getColumns()) {
                        c.setPrefWidth(w);
                    }
                };

        Runnable rebuildColumns =
                () -> {
                    double colW = 112;
                    try {
                        colW = Math.max(40, Double.parseDouble(colWidthField.getText().trim()));
                    } catch (NumberFormatException ignored) {
                    }
                    table.getColumns().clear();
                    for (int i = 0; i < headersRef.size(); i++) {
                        final int idx = i;
                        String title =
                                !headersRef.get(i).isBlank()
                                        ? headersRef.get(i)
                                        : ("\u5217" + (i + 1));
                        TableColumn<ObservableList<String>, String> col = new TableColumn<>(title);
                        col.setCellValueFactory(
                                cd -> {
                                    ObservableList<String> row = cd.getValue();
                                    String v =
                                            row != null && idx < row.size()
                                                    ? row.get(idx)
                                                    : "";
                                    return new javafx.beans.property.SimpleStringProperty(v);
                                });
                        col.setCellFactory(TextFieldTableCell.forTableColumn());
                        col.setOnEditCommit(
                                ev -> {
                                    ObservableList<String> row = ev.getRowValue();
                                    if (row == null) {
                                        return;
                                    }
                                    while (row.size() <= idx) {
                                        row.add("");
                                    }
                                    row.set(
                                            idx,
                                            ev.getNewValue() != null ? ev.getNewValue() : "");
                                    table.refresh();
                                });
                        col.setPrefWidth(colW);
                        table.getColumns().add(col);
                    }
                };

        Runnable applyLoaded =
                () -> {
                    rebuildColumns.run();
                    table.refresh();
                };

        Runnable syncFromEnv =
                () -> {
                    Map<String, String> env = envSupplier.get();
                    if (env != null) {
                        String p = trim(env.get(ENV_PM_AI_PLAN_INPUT_PATH));
                        if (!p.isEmpty() && pathField.getText().isBlank()) {
                            pathField.setText(p);
                        }
                        String sh = trim(env.get(ENV_TASK_PLAN_SHEET));
                        if (!sh.isEmpty()) {
                            sheetField.setText(sh);
                        }
                    }
                };

        Button browse =
                new Button("\u53c2\u7167\u2026");
        browse.setOnAction(
                e -> {
                    FileChooser ch = new FileChooser();
                    ch.setTitle("\u914d\u53f0\u8a08\u753b_\u30bf\u30b9\u30af\u5165\u529b ? \u30d5\u30a1\u30a4\u30eb");
                    ch.getExtensionFilters()
                            .addAll(
                                    new FileChooser.ExtensionFilter(
                                            "Tabular", "*.csv", "*.xlsx", "*.xlsm"),
                                    new FileChooser.ExtensionFilter("All", "*.*"));
                    File f = ch.showOpenDialog(owner);
                    if (f != null) {
                        pathField.setText(f.getAbsolutePath());
                    }
                });

        Button load =
                new Button("\u8aad\u8fbc");
        load.setOnAction(
                e -> {
                    syncFromEnv.run();
                    Path path = Path.of(pathField.getText().trim());
                    if (!java.nio.file.Files.isRegularFile(path)) {
                        log.accept("[plan-input] file not found: " + path);
                        return;
                    }
                    try {
                        PlanInputTabularIo.TabularSheet sheet =
                                PlanInputTabularIo.read(path, sheetField.getText().trim());
                        headersRef.clear();
                        headersRef.addAll(sheet.headers());
                        rows.clear();
                        for (List<String> line : sheet.rows()) {
                            ObservableList<String> r = FXCollections.observableArrayList(line);
                            while (r.size() < headersRef.size()) {
                                r.add("");
                            }
                            while (r.size() > headersRef.size()) {
                                r.remove(r.size() - 1);
                            }
                            rows.add(r);
                        }
                        applyLoaded.run();
                        log.accept(
                                "[plan-input] loaded rows="
                                        + rows.size()
                                        + " cols="
                                        + headersRef.size()
                                        + " path="
                                        + path);
                    } catch (Exception ex) {
                        log.accept("[plan-input] load error: " + ex.getMessage());
                    }
                });

        Button save =
                new Button("\u4fdd\u5b58");
        save.setOnAction(
                e -> {
                    Path path = Path.of(pathField.getText().trim());
                    if (pathField.getText().isBlank()) {
                        log.accept("[plan-input] save: path is empty");
                        return;
                    }
                    try {
                        List<List<String>> dataRows = new ArrayList<>();
                        for (ObservableList<String> r : rows) {
                            List<String> copy = new ArrayList<>(r);
                            while (copy.size() < headersRef.size()) {
                                copy.add("");
                            }
                            while (copy.size() > headersRef.size()) {
                                copy.remove(copy.size() - 1);
                            }
                            dataRows.add(copy);
                        }
                        PlanInputTabularIo.write(
                                path,
                                sheetField.getText().trim().isEmpty()
                                        ? DEFAULT_PLAN_INPUT_SHEET_NAME
                                        : sheetField.getText().trim(),
                                new PlanInputTabularIo.TabularSheet(headersRef, dataRows));
                        log.accept("[plan-input] saved " + path);
                    } catch (Exception ex) {
                        log.accept("[plan-input] save error: " + ex.getMessage());
                    }
                });

        Button addRow =
                new Button("\u884c\u8ffd\u52a0");
        addRow.setOnAction(
                e -> {
                    if (headersRef.isEmpty()) {
                        log.accept("[plan-input] load a file first (headers required)");
                        return;
                    }
                    ObservableList<String> r = FXCollections.observableArrayList();
                    for (int i = 0; i < headersRef.size(); i++) {
                        r.add("");
                    }
                    rows.add(r);
                });

        Button removeRows =
                new Button("\u884c\u524a\u9664");
        removeRows.setOnAction(
                e -> {
                    var sel = table.getSelectionModel().getSelectedItems();
                    if (sel.isEmpty()) {
                        return;
                    }
                    rows.removeAll(sel);
                    log.accept("[plan-input] removed " + sel.size() + " row(s)");
                });

        Label hint =
                new Label(
                        "PM_AI_PLAN_INPUT_PATH \u3068\u540c\u3058\u30d5\u30a1\u30a4\u30eb\u3092\u7de8\u96c6\u3057\u307e\u3059"
                                + " (\u6bb5\u968e2 load_planning_tasks_df)\u3002Excel \u306f"
                                + " \u30b7\u30fc\u30c8\u540d\u3092\u6307\u5b9a\u3002\u4fdd\u5b58\u6642"
                                + " .xlsx \u306f\u30c7\u30fc\u30bf\u306e\u307f\uff08\u30de\u30af\u30ed\u306f\u524a\u9664\u3055\u308c\u307e\u3059\uff09\u3002");

        hint.setWrapText(true);

        GridPaneAdapter gp = new GridPaneAdapter();
        gp.addRow(new Label("\u30d5\u30a1\u30a4\u30eb"), pathField, browse);
        gp.addRow(new Label("\u30b7\u30fc\u30c8\u540d"), sheetField, null);
        HBox actions = new HBox(8, load, save, addRow, removeRows);

        HBox planColStrip =
                new HBox(
                        8,
                        TableViewColumnSettingsStrip.create(table, applyDynamicColumnWidths, true),
                        new Label("\u65e2\u5b9a\u5217\u5e45(px)"),
                        colWidthField);
        planColStrip.setStyle("-fx-alignment: CENTER_LEFT;");

        VBox top = new VBox(8, gp.node(), actions, planColStrip, hint);
        top.setPadding(new Insets(8));

        BorderPane root = new BorderPane(table);
        root.setTop(top);
        BorderPane.setMargin(table, new Insets(0, 8, 8, 8));
        table.setMinHeight(240);

        javafx.application.Platform.runLater(
                () -> {
                    syncFromEnv.run();
                    if (!pathField.getText().isBlank()) {
                        load.fire();
                    }
                });

        return root;
    }

    private static String trim(String s) {
        return s != null ? s.trim() : "";
    }

    /** Minimal grid: label col + field col + optional third (button) column. */
    private static final class GridPaneAdapter {
        private final javafx.scene.layout.GridPane grid = new javafx.scene.layout.GridPane();
        private int nextRow;

        GridPaneAdapter() {
            grid.setHgap(8);
            grid.setVgap(8);
        }

        void addRow(Label lab, TextField field, Button extra) {
            int r = nextRow++;
            grid.add(lab, 0, r);
            grid.add(field, 1, r);
            if (extra != null) {
                grid.add(extra, 2, r);
            }
            javafx.scene.layout.GridPane.setHgrow(field, Priority.ALWAYS);
        }

        Parent node() {
            return grid;
        }
    }
}
