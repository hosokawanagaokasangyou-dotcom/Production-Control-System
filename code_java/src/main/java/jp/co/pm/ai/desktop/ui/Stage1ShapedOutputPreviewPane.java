package jp.co.pm.ai.desktop.ui;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicReference;
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
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;

/**
 * Read-only preview of the task-input table after {@code load_tasks_df} (header row, column alignment,
 * ???NO filter) ? {@code output/stage1_task_input_table.xlsx}, sheet {@link AppPaths#STAGE1_TASK_INPUT_PREVIEW_SHEET}.
 */
public final class Stage1ShapedOutputPreviewPane {

    /** Default sheet in {@link AppPaths#STAGE1_TASK_INPUT_PREVIEW_FILENAME}. */
    public static final String DEFAULT_STAGE1_PREVIEW_SHEET = AppPaths.STAGE1_TASK_INPUT_PREVIEW_SHEET;

    private Stage1ShapedOutputPreviewPane() {}

    public static Parent create(
            Stage owner,
            Supplier<Map<String, String>> envSupplier,
            Consumer<String> log) {
        return create(owner, envSupplier, log, null);
    }

    /**
     * @param registerReloadAfterStage1Success if non-null, receives a runnable to reload the preview
     *     xlsx when stage-1 Python exits 0.
     */
    public static Parent create(
            Stage owner,
            Supplier<Map<String, String>> envSupplier,
            Consumer<String> log,
            Consumer<Runnable> registerReloadAfterStage1Success) {

        TextField pathField = new TextField();
        pathField.setPromptText(
                "output/"
                        + AppPaths.STAGE1_TASK_INPUT_PREVIEW_FILENAME
                        + " (\u554f\u5408\u305b\u53d6\u8fbc\u5f8c\u30fb\u30bf\u30b9\u30af\u4e00\u89a7\u5316\u524d)");

        TextField sheetField = new TextField(DEFAULT_STAGE1_PREVIEW_SHEET);
        sheetField.setPromptText("Excel sheet name");

        TableView<ObservableList<String>> table = new TableView<>();
        table.setColumnResizePolicy(TableView.UNCONSTRAINED_RESIZE_POLICY);
        table.setEditable(false);
        table.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        table.setItems(rows);

        TextField colWidthField = new TextField("112");
        colWidthField.setMaxWidth(72);
        colWidthField.setPromptText("112");

        List<String> headersRef = new ArrayList<>();
        AtomicBoolean suppressColumnOrderPersistence = new AtomicBoolean(false);
        AtomicReference<List<TableColumnOrderPersistence.ColumnSpec>> persistedLayout =
                new AtomicReference<>(List.of());

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
                    suppressColumnOrderPersistence.set(true);
                    try {
                        double colW = 112;
                        try {
                            colW = Math.max(40, Double.parseDouble(colWidthField.getText().trim()));
                        } catch (NumberFormatException ignored) {
                        }
                        List<Double> widths =
                                TableColumnOrderPersistence.resolveWidthsForHeaders(
                                        headersRef, persistedLayout.get(), colW);
                        table.getColumns().clear();
                        for (int i = 0; i < headersRef.size(); i++) {
                            final int idx = i;
                            double prefW = i < widths.size() ? widths.get(i) : colW;
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
                            col.setCellFactory(TabularCellHighlight.stage1DateHighlightCellFactory(title));
                            col.setPrefWidth(prefW);
                            table.getColumns().add(col);
                        }
                    } finally {
                        suppressColumnOrderPersistence.set(false);
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
                    pathField.setText(AppPaths.defaultStage1TaskInputPreviewPath(env).toString());
                };

        Button fromEnv = new Button("\u74b0\u5883\u3088\u308a\u30d1\u30b9");
        fromEnv.setOnAction(e -> fillPathFromEnv.run());

        Button browse = new Button("\u53c2\u7167\u2026");
        browse.setOnAction(
                e -> {
                    FileChooser ch = new FileChooser();
                    ch.setTitle(AppPaths.STAGE1_TASK_INPUT_PREVIEW_FILENAME);
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
                        sheet = DEFAULT_STAGE1_PREVIEW_SHEET;
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
                        List<TableColumnOrderPersistence.ColumnSpec> lay =
                                TableColumnOrderPersistence.loadLayout(
                                        TableColumnOrderPersistence.TableId.STAGE1_PREVIEW);
                        persistedLayout.set(lay);
                        TableColumnOrderPersistence.applyLogicalColumnOrder(
                                headersRef,
                                rows,
                                lay.stream().map(TableColumnOrderPersistence.ColumnSpec::title).toList());
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
                        "\u554f\u5408\u305b xlsx \u3092\u8aad\u307f\u8fbc\u307f\u3001\u30d8\u30c3\u30c0\u30fc\u884c\u3068"
                                + "\u5217\u540d\u3092\u6574\u3048\u305f\u76f4\u5f8c\uff08\u4f9d\u983cNO"
                                + "\u304c\u3042\u308b\u884c\u306e\u307f\uff09\u3002"
                                + " \u30de\u30b9\u30bf\u30fb\u914d\u53f0\u8a66\u884c\u9806\u4ed8\u4e0e\u524d\u306e"
                                + " stage1_task_input_table.xlsx"
                                + " \u30b7\u30fc\u30c8\u300c"
                                + DEFAULT_STAGE1_PREVIEW_SHEET
                                + "\u300d\u3092\u8868\u793a\u3057\u307e\u3059\u3002");
        hint.setWrapText(true);

        HBox pathRow = new HBox(8, new Label("\u30d5\u30a1\u30a4\u30eb"), pathField, fromEnv, browse);
        HBox.setHgrow(pathField, Priority.ALWAYS);
        HBox sheetRow = new HBox(8, new Label("\u30b7\u30fc\u30c8"), sheetField);
        HBox.setHgrow(sheetField, Priority.ALWAYS);
        HBox actions = new HBox(8, load);
        HBox colStrip =
                new HBox(
                        8,
                        TableViewColumnSettingsStrip.create(table, applyDynamicColumnWidths, false),
                        new Label("\u5217\u5e45(px)"),
                        colWidthField);
        colStrip.setStyle("-fx-alignment: CENTER_LEFT;");

        VBox top = new VBox(8, hint, pathRow, sheetRow, actions, colStrip);
        top.setPadding(new Insets(8));

        VBox root = new VBox(8, top, table);
        root.setFillWidth(true);
        VBox.setVgrow(table, Priority.ALWAYS);
        VBox.setMargin(table, new Insets(0, 8, 8, 8));
        table.setMinHeight(240);

        if (registerReloadAfterStage1Success != null) {
            registerReloadAfterStage1Success.accept(
                    () -> {
                        fillPathFromEnv.run();
                        sheetField.setText(DEFAULT_STAGE1_PREVIEW_SHEET);
                        load.fire();
                    });
        }

        TableColumnOrderPersistence.installColumnLayoutWatcher(
                table,
                TableColumnOrderPersistence.TableId.STAGE1_PREVIEW,
                suppressColumnOrderPersistence::get);

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
