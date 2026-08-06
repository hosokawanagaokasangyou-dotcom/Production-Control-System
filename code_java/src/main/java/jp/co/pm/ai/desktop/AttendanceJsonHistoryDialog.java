package jp.co.pm.ai.desktop;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.function.Consumer;

import com.fasterxml.jackson.databind.JsonNode;

import javafx.collections.FXCollections;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Label;
import javafx.scene.control.ListCell;
import javafx.scene.control.ListView;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.StageStyle;

import jp.co.pm.ai.desktop.config.AppPaths;

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

        String displayText() {
            String head = label != null && !label.isBlank() ? label : id;
            return head
                    + "  会社="
                    + companyCalendarRevision
                    + " メンバー="
                    + memberAttendanceRevision
                    + "  "
                    + (savedAt != null ? savedAt : "");
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

        ListView<HistoryEntry> listView = new ListView<>();
        listView.setCellFactory(
                lv ->
                        new ListCell<>() {
                            @Override
                            protected void updateItem(HistoryEntry item, boolean empty) {
                                super.updateItem(item, empty);
                                setText(empty || item == null ? null : item.displayText());
                            }
                        });
        listView.setPrefHeight(280);

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
                                listView.setItems(FXCollections.observableArrayList(items));
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
                    HistoryEntry sel = listView.getSelectionModel().getSelectedItem();
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
                        listView,
                        status,
                        spacer,
                        buttons);
        root.setPadding(new Insets(16));
        root.setPrefWidth(560);
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
