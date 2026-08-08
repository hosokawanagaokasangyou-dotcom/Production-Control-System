package jp.co.pm.ai.desktop;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import com.fasterxml.jackson.databind.JsonNode;

import javafx.collections.FXCollections;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.beans.property.SimpleIntegerProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.scene.control.Label;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.StageStyle;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.ui.JapanDateTimeDisplay;

/**
 * 勤怠正本 JSON（attendance-data.json）の世代一覧表示と復元。
 */
public final class AttendanceJsonHistoryDialog {

    private AttendanceJsonHistoryDialog() {}

    public record HistoryEntry(
            String id,
            String label,
            String savedAt,
            int companyCalendarRevision,
            int memberAttendanceRevision) {

        String savedAtDisplay() {
            return JapanDateTimeDisplay.formatSavedAtForDisplay(savedAt);
        }

        String displayText() {
            String head = label != null && !label.isBlank() ? label : id;
            return head
                    + "  会社="
                    + companyCalendarRevision
                    + " メンバー="
                    + memberAttendanceRevision
                    + "  "
                    + savedAtDisplay();
        }
    }

    public static void show(MainShellController shell, Runnable onRestored) {
        if (shell == null) {
            return;
        }
        Stage stage = new Stage(StageStyle.DECORATED);
        stage.initOwner(shell.primaryStageForDialogs());
        stage.initModality(Modality.WINDOW_MODAL);
        stage.setTitle("勤怠 JSON 世代から復元");

        Label status = new Label("世代一覧を読み込み中…");
        status.setWrapText(true);
        Label pathLabel = new Label();
        pathLabel.setWrapText(true);

        TableView<HistoryEntry> tableView = new TableView<>();
        tableView.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_FLEX_LAST_COLUMN);
        TableColumn<HistoryEntry, String> labelCol = new TableColumn<>("操作");
        labelCol.setCellValueFactory(
                cd -> {
                    HistoryEntry e = cd.getValue();
                    String head =
                            e != null && e.label() != null && !e.label().isBlank()
                                    ? e.label()
                                    : e != null ? e.id() : "";
                    return new SimpleStringProperty(head);
                });
        labelCol.setMinWidth(140);
        TableColumn<HistoryEntry, Integer> companyCol = new TableColumn<>("会社 rev");
        companyCol.setCellValueFactory(
                cd ->
                        new SimpleIntegerProperty(
                                cd.getValue() != null
                                        ? cd.getValue().companyCalendarRevision()
                                        : 0)
                                .asObject());
        companyCol.setMinWidth(72);
        TableColumn<HistoryEntry, Integer> memberCol = new TableColumn<>("メンバー rev");
        memberCol.setCellValueFactory(
                cd ->
                        new SimpleIntegerProperty(
                                cd.getValue() != null
                                        ? cd.getValue().memberAttendanceRevision()
                                        : 0)
                                .asObject());
        memberCol.setMinWidth(88);
        TableColumn<HistoryEntry, String> savedAtCol = new TableColumn<>("保存日時");
        savedAtCol.setCellValueFactory(
                cd ->
                        new SimpleStringProperty(
                                cd.getValue() != null ? cd.getValue().savedAtDisplay() : ""));
        savedAtCol.setMinWidth(150);
        tableView.getColumns().addAll(labelCol, companyCol, memberCol, savedAtCol);
        tableView.setMinHeight(220);
        VBox.setVgrow(tableView, Priority.ALWAYS);

        Button refreshBtn = new Button("一覧更新");
        Button restoreBtn = new Button("選択世代を復元");
        restoreBtn.setDefaultButton(true);
        Button closeBtn = new Button("閉じる");

        Runnable loadList =
                () -> {
                    status.setText("世代一覧を読み込み中…");
                    restoreBtn.setDisable(true);
                    shell.runAttendanceDataIoAsync(
                            shell.buildAttendanceDataIoRequest("history_list"),
                            node -> {
                                pathLabel.setText(
                                        "世代フォルダ: "
                                                + node.path("history_dir").asText("")
                                                + "　保持上限 "
                                                + node.path("max_entries").asInt(
                                                        AppPaths.ATTENDANCE_JSON_HISTORY_MAX_GENERATIONS)
                                                + " 世代");
                                List<HistoryEntry> items = parseEntries(node.path("entries"));
                                tableView.setItems(FXCollections.observableArrayList(items));
                                if (!items.isEmpty()) {
                                    tableView.getSelectionModel().select(0);
                                }
                                status.setText(
                                        items.isEmpty()
                                                ? "保存済み世代がありません。"
                                                : "復元する世代を選んでください。");
                                restoreBtn.setDisable(items.isEmpty());
                            },
                            err -> status.setText("一覧取得失敗: " + err));
                };

        refreshBtn.setOnAction(e -> loadList.run());
        loadList.run();

        restoreBtn.setOnAction(
                e -> {
                    HistoryEntry sel = tableView.getSelectionModel().getSelectedItem();
                    if (sel == null) {
                        status.setText("復元する世代を一覧から選んでください。");
                        return;
                    }
                    Alert confirm = new Alert(AlertType.CONFIRMATION);
                    confirm.setTitle("復元の確認");
                    confirm.setHeaderText(null);
                    confirm.setContentText(
                            "選択した世代（"
                                    + sel.displayText()
                                    + "）で勤怠正本（attendance-data.json）を上書きします。\n"
                                    + "復元前の現行ファイルは自動で世代退避されます。\n\n続行しますか？");
                    confirm.initOwner(stage);
                    Optional<ButtonType> ans = confirm.showAndWait();
                    if (ans.isEmpty() || ans.get() != ButtonType.OK) {
                        return;
                    }
                    status.setText("復元中…");
                    restoreBtn.setDisable(true);
                    shell.runAttendanceDataIoAsync(
                            shell.buildAttendanceDataIoRequest(
                                    "history_restore", sel.id()),
                            node -> {
                                status.setText("復元完了: " + sel.label());
                                shell.appendLog(
                                        "[attendance-history] restored: " + sel.id());
                                shell.refreshAttendanceReadiness();
                                if (onRestored != null) {
                                    onRestored.run();
                                }
                                loadList.run();
                                Alert done = new Alert(AlertType.INFORMATION);
                                done.setTitle("復元完了");
                                done.setHeaderText(null);
                                done.setContentText("勤怠 JSON を選択世代に復元しました。");
                                done.initOwner(stage);
                                done.showAndWait();
                            },
                            err -> {
                                status.setText("復元失敗: " + err);
                                restoreBtn.setDisable(false);
                            });
                });

        closeBtn.setOnAction(e -> stage.close());

        HBox buttons = new HBox(8, refreshBtn, restoreBtn, closeBtn);
        buttons.setAlignment(Pos.CENTER_LEFT);
        Region spacer = new Region();
        VBox.setVgrow(spacer, Priority.ALWAYS);
        VBox root =
                new VBox(
                        10,
                        new Label("会社カレンダー・メンバー勤怠の JSON 正本を過去世代から復元します。"),
                        pathLabel,
                        tableView,
                        status,
                        spacer,
                        buttons);
        root.setPadding(new Insets(16));
        root.setPrefWidth(640);
        root.setPrefHeight(420);
        stage.setMinWidth(520);
        stage.setMinHeight(360);
        stage.setScene(new javafx.scene.Scene(root));
        stage.showAndWait();
    }

    private static List<HistoryEntry> parseEntries(JsonNode arr) {
        List<HistoryEntry> out = new ArrayList<>();
        if (arr == null || !arr.isArray()) {
            return out;
        }
        for (JsonNode n : arr) {
            String id = n.path("id").asText("");
            if (id.isBlank()) {
                continue;
            }
            out.add(
                    new HistoryEntry(
                            id,
                            n.path("label").asText(""),
                            n.path("savedAt").asText(""),
                            n.path("company_calendar_revision").asInt(0),
                            n.path("member_attendance_revision").asInt(0)));
        }
        return out;
    }
}
