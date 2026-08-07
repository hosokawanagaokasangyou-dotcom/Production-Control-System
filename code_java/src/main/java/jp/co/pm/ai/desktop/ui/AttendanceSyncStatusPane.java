package jp.co.pm.ai.desktop.ui;

import com.fasterxml.jackson.databind.JsonNode;

import javafx.geometry.Insets;
import javafx.scene.control.Label;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.VBox;

/** 勤怠 JSON / 勤怠カレンダー.xlsx / 段階2準備状態の簡易パネル。 */
public final class AttendanceSyncStatusPane extends VBox {

    private final Label jsonLabel = new Label();
    private final Label calendarXlsxLabel = new Label();
    private final Label readinessLabel = new Label();

    public AttendanceSyncStatusPane() {
        getStyleClass().add("pm-attendance-sync-status");
        setSpacing(4);
        setPadding(new Insets(6, 8, 6, 8));
        GridPane grid = new GridPane();
        grid.setHgap(12);
        grid.setVgap(4);
        grid.addRow(0, new Label("JSON正本:"), jsonLabel);
        grid.addRow(1, new Label("勤怠カレンダー:"), calendarXlsxLabel);
        grid.addRow(2, new Label("段階2準備:"), readinessLabel);
        getChildren().add(grid);
    }

    public void updateFromReadiness(JsonNode node) {
        if (node == null) {
            jsonLabel.setText("—");
            calendarXlsxLabel.setText("—");
            readinessLabel.setText("—");
            return;
        }
        jsonLabel.setText(
                node.path("json_path").asText("")
                        + (node.path("json_exists").asBoolean(false) ? " ✓" : " 未作成"));
        String calendarPath = node.path("calendar_xlsx_path").asText("");
        boolean calendarExists = node.path("calendar_xlsx_exists").asBoolean(false);
        String calendarAt = node.path("calendar_xlsx_export_at").asText("");
        String calendarLine =
                calendarPath
                        + (calendarExists ? " ✓" : " 未作成")
                        + (calendarAt.isBlank() ? "" : " / 出力 " + calendarAt);
        calendarXlsxLabel.setText(calendarLine);
        boolean ready = node.path("stage2_ready").asBoolean(false);
        String issues = "";
        if (node.path("issues").isArray() && node.path("issues").size() > 0) {
            issues = node.path("issues").get(0).asText("");
        }
        readinessLabel.setText(
                ready ? "OK（メンバー勤怠 " + node.path("member_cells_in_month").asInt(0) + " セル）"
                        : (issues.isBlank() ? "未準備" : issues));
        readinessLabel.getStyleClass().remove("pm-attendance-sync-warn");
        if (!ready) {
            readinessLabel.getStyleClass().add("pm-attendance-sync-warn");
        }
    }
}
