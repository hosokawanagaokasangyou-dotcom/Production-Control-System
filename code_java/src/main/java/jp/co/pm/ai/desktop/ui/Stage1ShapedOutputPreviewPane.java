package jp.co.pm.ai.desktop.ui;

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

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;

/**
 * Read-only preview of stage-1 shaped output: {@code output/plan_input_tasks.xlsx}, sheet
 * {@code \u30bf\u30b9\u30af\u4e00\u89a7} (same as {@code planning_core.run_stage1_extract}).
 */
public final class Stage1ShapedOutputPreviewPane {

    /** Sheet name written by {@code run_stage1_extract} ({@code to_excel(..., sheet_name="..." )}). */
    public static final String DEFAULT_STAGE1_OUTPUT_SHEET = "\u30bf\u30b9\u30af\u4e00\u89a7";

    private Stage1ShapedOutputPreviewPane() {}

    public static Parent create(
            Stage owner,
            Supplier<Map<String, String>> envSupplier,
            Consumer<String> log) {

        TextField pathField = new TextField();
        pathField.setPromptText(
                "output/"
                        + AppPaths.STAGE1_PLAN_TASKS_FILENAME
                        + " (\u6bb5\u968e1\u5b8c\u4e86\u5f8c)");

        TextField sheetField = new TextField(DEFAULT_STAGE1_OUTPUT_SHEET);
        sheetField.setPromptText("Excel sheet name");

        TableView<ObservableList<String>> table = new TableView<>();
        table.setEditable(false);
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
                        col.setEditable(false);
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
                        col.setPrefWidth(colW);
                        table.getColumns().add(col);
                    }
                };

        Runnable applyLoaded =
                () -> {
                    rebuildColumns.run();
                    table.refresh();
                };

        Runnable fillPathFromEnv =
                () -> {
                    Map<String, String> env = envSupplier.get();
                    if (env == null) {
                        return;
                    }
                    pathField.setText(AppPaths.defaultStage1PlanTasksPath(env).toString());
                };

        Button fromEnv = new Button("\u74b0\u5883\u3088\u308a\u30d1\u30b9");
        fromEnv.setOnAction(e -> fillPathFromEnv.run());

        Button browse = new Button("\u53c2\u7167\u2026");
        browse.setOnAction(
                e -> {
                    FileChooser ch = new FileChooser();
                    ch.setTitle("plan_input_tasks.xlsx");
                    ch.getExtensionFilters()
                            .addAll(
                                    new FileChooser.ExtensionFilter("Excel", "*.xlsx", "*.xlsm"),
                                    new FileChooser.ExtensionFilter("All", "*.*"));
                    String cur = pathField.getText();
                    if (cur != null && !cur.isBlank()) {
                        try {
                            Path p = Path.of(cur.trim());
                            if (java.nio.file.Files.isRegularFile(p) && p.getParent() != null) {
                                ch.setInitialDirectory(p.getParent().toFile());
                            }
                        } catch (Exception ignored) {
                        }
                    }
                    var f = ch.showOpenDialog(owner);
                    if (f != null) {
                        pathField.setText(f.getAbsolutePath());
                    }
                });

        Button load = new Button("\u8aad\u8fbc");
        load.setOnAction(
                e -> {
                    if (pathField.getText().isBlank()) {
                        fillPathFromEnv.run();
                    }
                    Path path = Path.of(pathField.getText().trim());
                    if (!java.nio.file.Files.isRegularFile(path)) {
                        log.accept("[stage1-preview] file not found: " + path);
                        return;
                    }
                    String sheet = sheetField.getText().trim();
                    if (sheet.isEmpty()) {
                        sheet = DEFAULT_STAGE1_OUTPUT_SHEET;
                    }
                    try {
                        PlanInputTabularIo.TabularSheet sh = PlanInputTabularIo.read(path, sheet);
                        headersRef.clear();
                        headersRef.addAll(sh.headers());
                        rows.clear();
                        for (List<String> line : sh.rows()) {
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
                                "[stage1-preview] loaded rows="
                                        + rows.size()
                                        + " cols="
                                        + headersRef.size()
                                        + " path="
                                        + path);
                    } catch (Exception ex) {
                        log.accept("[stage1-preview] load error: " + ex.getMessage());
                    }
                });

        Label hint =
                new Label(
                        "\u6bb5\u968e1 (run_stage1_extract) \u5f8c\u306e\u6210\u5f62\u7d50\u679c"
                                + " \u2014 \u901a\u5e38"
                                + " code/python/output/plan_input_tasks.xlsx"
                                + " \u306e\u30b7\u30fc\u30c8"
                                + " \u300c"
                                + DEFAULT_STAGE1_OUTPUT_SHEET
                                + "\u300d"
                                + "\u3002"
                                + " \u9664\u5916\u30eb\u30fc\u30eb\u900f\u904e\u30fb\u914d\u53f0\u8a66\u884c\u9806"
                                + "\u306a\u3069\u6e21\u3057\u5f8c\u306e\u30c6\u30fc\u30d6\u30eb\u3067\u3059\uff08"
                                + "\u8aad\u307f\u53d6\u308a\u5c02\u7528\uff09\u3002");
        hint.setWrapText(true);

        HBox pathRow = new HBox(8, new Label("\u30d5\u30a1\u30a4\u30eb"), pathField, fromEnv, browse);
        HBox.setHgrow(pathField, Priority.ALWAYS);
        HBox sheetRow = new HBox(8, new Label("\u30b7\u30fc\u30c8"), sheetField);
        HBox.setHgrow(sheetField, Priority.ALWAYS);
        HBox actions = new HBox(8, load);
        HBox colStrip =
                new HBox(
                        8,
                        TableViewColumnSettingsStrip.create(table, applyDynamicColumnWidths, true),
                        new Label("\u5217\u5e45(px)"),
                        colWidthField);
        colStrip.setStyle("-fx-alignment: CENTER_LEFT;");

        VBox top = new VBox(8, hint, pathRow, sheetRow, actions, colStrip);
        top.setPadding(new Insets(8));

        BorderPane root = new BorderPane(table);
        root.setTop(top);
        BorderPane.setMargin(table, new Insets(0, 8, 8, 8));
        table.setMinHeight(240);
        VBox.setVgrow(table, Priority.ALWAYS);

        javafx.application.Platform.runLater(
                () -> {
                    fillPathFromEnv.run();
                    if (!pathField.getText().isBlank()) {
                        load.fire();
                    }
                });

        return root;
    }
}
