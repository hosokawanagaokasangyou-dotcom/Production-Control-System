package jp.co.pm.ai.desktop.ui;

import com.fasterxml.jackson.databind.JsonNode;

import javafx.geometry.Insets;
import javafx.scene.control.Label;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.VBox;

/** 勤怠 JSON / 表示用 Excel / master 出力状態の簡易パネル。 */
public final class AttendanceSyncStatusPane extends VBox {

    private final Label jsonLabel = new Label();
    private final Label viewLabel = new Label();
    private final Label masterLabel = new Label();
    private final Label readinessLabel = new Label();

    public AttendanceSyncStatusPane() {
        getStyleClass().add("pm-attendance-sync-status");
        setSpacing(4);
        setPadding(new Insets(6, 8, 6, 8));
        GridPane grid = new GridPane();
        grid.setHgap(12);
        grid.setVgap(4);
        grid.addRow(0, new Label("JSON正本:"), jsonLabel);
        grid.addRow(1, new Label("閲覧用xlsx:"), viewLabel);
        grid.addRow(2, new Label("master出力:"), masterLabel);
        grid.addRow(3, new Label("段階2準備:"), readinessLabel);
        getChildren().add(grid);
    }

    public void updateFromReadiness(JsonNode node) {
        if (node == null) {
            jsonLabel.setText("—");
            viewLabel.setText("—");
            masterLabel.setText("—");
            readinessLabel.setText("—");
            return;
        }
        jsonLabel.setText(
                node.path("json_path").asText("")
                        + (node.path("json_exists").asBoolean(false) ? " ✓" : " 未作成"));
        viewLabel.setText(
                node.path("view_xlsx_path").asText("")
                        + (node.path("view_xlsx_exists").asBoolean(false) ? " ✓" : " 未生成"));
        String masterAt = node.path("master_export_at").asText("");
        masterLabel.setText(masterAt.isBlank() ? "未出力" : masterAt);
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
