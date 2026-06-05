package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;
import java.util.function.Supplier;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.Tooltip;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.io.PostProcessingProductMasterIo;

/**
 * アラジンマスタフォルダ内の xlsx を表形式で参照表示する。
 */
public final class PostProcessingExcelMasterCatalogPane {

    private PostProcessingExcelMasterCatalogPane() {}

    public static VBox build(
            Supplier<Map<String, String>> uiEnv,
            String fileName,
            String panelTitle,
            String note,
            Runnable invalidateBeforeReload,
            Consumer<String> log) {
        Consumer<String> logFn = log != null ? log : s -> {};

        Label statusLabel = new Label("読込中...");
        statusLabel.setWrapText(true);
        statusLabel.setStyle("-fx-font-size: 11px; -fx-font-weight: bold;");

        Label sourceLabel = new Label();
        sourceLabel.setWrapText(true);
        sourceLabel.setStyle("-fx-font-size: 10px;");

        ObservableList<Map<String, String>> rows = FXCollections.observableArrayList();
        TableView<Map<String, String>> table = new TableView<>(rows);
        table.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY);
        table.setPlaceholder(new Label("データがありません。「再読込」でマスタを読み直してください。"));
        VBox.setVgrow(table, Priority.ALWAYS);
        table.setMinHeight(240);
        table.setPrefHeight(420);

        Runnable reload =
                () -> {
                    statusLabel.setText("マスタを読込中...");
                    Thread t =
                            new Thread(
                                    () -> {
                                        try {
                                            if (invalidateBeforeReload != null) {
                                                invalidateBeforeReload.run();
                                            }
                                            Map<String, String> ui = uiEnv.get();
                                            Path path =
                                                    AppPaths.resolveAladdinMasterDir(ui)
                                                            .resolve(fileName);
                                            LoadedSheet loaded = loadSheet(path);
                                            Platform.runLater(
                                                    () ->
                                                            applyLoaded(
                                                                    loaded,
                                                                    statusLabel,
                                                                    sourceLabel,
                                                                    table,
                                                                    rows));
                                        } catch (Exception ex) {
                                            Platform.runLater(
                                                    () -> {
                                                        rows.clear();
                                                        table.getColumns().clear();
                                                        statusLabel.setText(
                                                                "読込失敗: " + ex.getMessage());
                                                        sourceLabel.setText("");
                                                    });
                                            logFn.accept(
                                                    "[excel-master-catalog] "
                                                            + fileName
                                                            + ": "
                                                            + ex.getMessage());
                                        }
                                    },
                                    "excel-master-catalog-" + fileName);
                    t.setDaemon(true);
                    t.start();
                };

        Button btnReload = new Button("再読込");
        btnReload.getStyleClass().add("btn-reload");
        btnReload.setTooltip(
                new Tooltip(
                        "PM_AI_ALADDIN_MASTER_DIR（サマリ Excel と同じフォルダ）から "
                                + fileName
                                + " を再読込"));
        btnReload.setOnAction(e -> reload.run());

        Label title = new Label(panelTitle);
        title.getStyleClass().add("paper-main-title");

        Label noteLabel = new Label(note != null ? note : "");
        noteLabel.setWrapText(true);
        noteLabel.setStyle("-fx-font-size: 10px;");
        noteLabel.setManaged(note != null && !note.isBlank());
        noteLabel.setVisible(note != null && !note.isBlank());

        HBox actions = new HBox(8, btnReload);
        actions.setAlignment(Pos.CENTER_LEFT);

        VBox root = new VBox(10);
        root.getStyleClass().add("form-tab-container");
        root.setFillWidth(true);
        root.setPadding(new Insets(8, 12, 12, 12));
        root.getChildren()
                .addAll(title, noteLabel, actions, statusLabel, sourceLabel, table);

        reload.run();
        return root;
    }

    private record LoadedSheet(Path path, List<String> headers, List<Map<String, String>> rows) {}

    private static LoadedSheet loadSheet(Path path) throws IOException {
        if (!Files.isRegularFile(path)) {
            return new LoadedSheet(path, List.of(), List.of());
        }
        PlanInputTabularIo.TabularSheet sheet =
                PlanInputTabularIo.read(path, PostProcessingProductMasterIo.DEFAULT_SHEET_NAME);
        List<String> headers = List.copyOf(sheet.headers());
        List<Map<String, String>> dataRows = new ArrayList<>();
        for (List<String> row : sheet.rows()) {
            Map<String, String> map =
                    new LinkedHashMap<>(
                            PostProcessingProductMasterIo.rowToMap(headers, row));
            dataRows.add(map);
        }
        return new LoadedSheet(path, headers, dataRows);
    }

    private static void applyLoaded(
            LoadedSheet loaded,
            Label statusLabel,
            Label sourceLabel,
            TableView<Map<String, String>> table,
            ObservableList<Map<String, String>> rows) {
        rebuildColumns(table, loaded.headers());
        rows.setAll(loaded.rows());
        if (!Files.isRegularFile(loaded.path())) {
            statusLabel.setText("マスタファイルが見つかりません: " + loaded.path());
            sourceLabel.setText("");
            return;
        }
        statusLabel.setText("登録 " + loaded.rows().size() + " 件");
        sourceLabel.setText("参照ファイル: " + loaded.path().toAbsolutePath().normalize());
    }

    private static void rebuildColumns(TableView<Map<String, String>> table, List<String> headers) {
        table.getColumns().clear();
        if (headers.isEmpty()) {
            return;
        }
        for (String header : headers) {
            String title = header != null && !header.isBlank() ? header : "列";
            TableColumn<Map<String, String>, String> col = new TableColumn<>(title);
            col.setCellValueFactory(
                    cd -> {
                        Map<String, String> row = cd.getValue();
                        String v = row != null ? row.getOrDefault(header, "") : "";
                        return new javafx.beans.property.SimpleStringProperty(
                                v != null ? v : "");
                    });
            col.setPrefWidth(Math.min(Math.max(title.length() * 10 + 40, 90), 220));
            table.getColumns().add(col);
        }
    }
}
