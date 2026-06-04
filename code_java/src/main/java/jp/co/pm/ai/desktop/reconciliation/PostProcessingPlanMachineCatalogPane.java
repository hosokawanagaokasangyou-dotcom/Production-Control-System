package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
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

/**
 * 加工計画DATA 由来の機械コード・機械名一覧（参照用）。
 */
public final class PostProcessingPlanMachineCatalogPane {

    public record MachineRow(String machineCode, String machineName) {}

    private PostProcessingPlanMachineCatalogPane() {}

    public static VBox build(Supplier<Map<String, String>> uiEnv, Consumer<String> log) {
        Consumer<String> logFn = log != null ? log : s -> {};

        Label statusLabel = new Label("読込中...");
        statusLabel.setWrapText(true);
        statusLabel.setStyle("-fx-font-size: 11px; -fx-font-weight: bold;");

        Label sourceLabel = new Label();
        sourceLabel.setWrapText(true);
        sourceLabel.setStyle("-fx-font-size: 10px;");

        Label metaLabel = new Label();
        metaLabel.setWrapText(true);
        metaLabel.setStyle("-fx-font-size: 10px;");

        ObservableList<MachineRow> rows = FXCollections.observableArrayList();
        TableView<MachineRow> table = new TableView<>(rows);
        table.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY);
        table.setPlaceholder(new Label("データがありません。「再読込」で加工計画を読み直してください。"));

        TableColumn<MachineRow, String> colCode = new TableColumn<>("機械");
        colCode.setCellValueFactory(cd -> new javafx.beans.property.SimpleStringProperty(
                cd.getValue() != null ? cd.getValue().machineCode() : ""));
        colCode.setPrefWidth(120);

        TableColumn<MachineRow, String> colName = new TableColumn<>("機械名");
        colName.setCellValueFactory(cd -> new javafx.beans.property.SimpleStringProperty(
                cd.getValue() != null ? cd.getValue().machineName() : ""));
        colName.setPrefWidth(280);

        table.getColumns().addAll(colCode, colName);
        VBox.setVgrow(table, Priority.ALWAYS);
        table.setMinHeight(240);
        table.setPrefHeight(420);

        Runnable reload =
                () -> {
                    statusLabel.setText("加工計画を読込中...");
                    Thread t =
                            new Thread(
                                    () -> {
                                        try {
                                            PostProcessingPlanMachineLookup.invalidate();
                                            PostProcessingPlanMachineLookup.Snapshot snap =
                                                    PostProcessingPlanMachineLookup.snapshot(
                                                            uiEnv.get());
                                            List<MachineRow> loaded = toRows(snap);
                                            Platform.runLater(
                                                    () -> applySnapshot(
                                                            snap,
                                                            loaded,
                                                            statusLabel,
                                                            sourceLabel,
                                                            metaLabel,
                                                            rows));
                                        } catch (Exception ex) {
                                            Platform.runLater(
                                                    () -> {
                                                        rows.clear();
                                                        statusLabel.setText(
                                                                "読込失敗: " + ex.getMessage());
                                                        sourceLabel.setText("");
                                                        metaLabel.setText("");
                                                    });
                                            logFn.accept(
                                                    "[plan-machine-db] "
                                                            + ex.getMessage());
                                        }
                                    },
                                    "plan-machine-catalog-reload");
                    t.setDaemon(true);
                    t.start();
                };

        Button btnReload = new Button("再読込");
        btnReload.getStyleClass().add("btn-reload");
        btnReload.setTooltip(
                new Tooltip(
                        "PM_AI_TASK_INPUT_SOURCE_DIR 最新 / PM_AI_PROCESSING_PLAN_PATH から再構築"));
        btnReload.setOnAction(e -> reload.run());

        Label title = new Label("加工計画・機械コード一覧");
        title.getStyleClass().add("paper-main-title");

        Label note =
                new Label(
                        "加工計画DATA の「機械」「機械名」列から一意一覧を表示します。"
                                + " 後加工商品マスタ編集の機械コード1〜8コンボと同じデータです。");
        note.setWrapText(true);
        note.setStyle("-fx-font-size: 10px;");

        HBox actions = new HBox(8, btnReload);
        actions.setAlignment(Pos.CENTER_LEFT);

        VBox root = new VBox(10);
        root.getStyleClass().add("form-tab-container");
        root.setFillWidth(true);
        root.setPadding(new Insets(8, 12, 12, 12));
        root.getChildren()
                .addAll(title, note, actions, statusLabel, sourceLabel, metaLabel, table);
        VBox.setVgrow(table, Priority.ALWAYS);

        reload.run();
        return root;
    }

    private static List<MachineRow> toRows(PostProcessingPlanMachineLookup.Snapshot snap) {
        List<MachineRow> out = new ArrayList<>();
        if (snap == null || snap.machineCodeToName().isEmpty()) {
            return out;
        }
        for (Map.Entry<String, String> e : snap.machineCodeToName().entrySet()) {
            out.add(new MachineRow(e.getKey(), e.getValue() != null ? e.getValue() : ""));
        }
        return out;
    }

    private static void applySnapshot(
            PostProcessingPlanMachineLookup.Snapshot snap,
            List<MachineRow> loaded,
            Label statusLabel,
            Label sourceLabel,
            Label metaLabel,
            ObservableList<MachineRow> rows) {
        rows.setAll(loaded);
        if (snap == null || !snap.loaded()) {
            statusLabel.setText("機械一覧は空です（加工計画に機械／機械名が無いか未読込）。");
            sourceLabel.setText("");
            metaLabel.setText("");
            return;
        }
        String pathText =
                snap.sourcePath() != null && Files.isRegularFile(snap.sourcePath())
                        ? snap.sourcePath().toString()
                        : "(不明)";
        statusLabel.setText("登録 " + loaded.size() + " 件");
        sourceLabel.setText("参照ファイル: " + pathText);
        metaLabel.setText(
                "列: 機械="
                        + (snap.hasCodeColumn() ? "あり" : "なし")
                        + "、機械名="
                        + (snap.hasNameColumn() ? "あり" : "なし")
                        + (snap.hasCodeColumn() && !snap.hasNameColumn()
                                ? "（機械未設定行は機械名をキーにしています）"
                                : ""));
    }
}
