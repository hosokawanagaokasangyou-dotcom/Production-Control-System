package jp.co.pm.ai.desktop.dispatch.rules.ui;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Label;
import javafx.scene.layout.FlowPane;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.shape.Circle;

/** Roll flow: 投入前原反数 → 接続 → SEC前仕掛かり → SEC完了. */
public final class DispatchRuleRollStageVisualPane extends GridPane {

    private static final int MAX_ICONS = 24;
    private static final double ICON_RADIUS = 6.5;

    private final StageColumn preInputRaw = new StageColumn("投入前原反数", Color.web("#95A5A6"));
    private final StageColumn connection = new StageColumn("接続", Color.web("#E67E22"));
    private final StageColumn secBeforeWip = new StageColumn("SEC前仕掛かり", Color.web("#9B59B6"));
    private final StageColumn secComplete = new StageColumn("SEC完了", Color.web("#27AE60"));

    public DispatchRuleRollStageVisualPane() {
        setHgap(12);
        setVgap(4);
        setPadding(new Insets(4, 0, 4, 0));
        add(preInputRaw.root, 0, 0);
        add(connection.root, 1, 0);
        add(secBeforeWip.root, 2, 0);
        add(secComplete.root, 3, 0);
        for (int c = 0; c < 4; c++) {
            GridPane.setHgrow(getChildren().get(c), javafx.scene.layout.Priority.ALWAYS);
        }
        clear();
    }

    public void clear() {
        update(0, 0, 0, 0, 0, 20, 0, 0);
    }

    public void update(
            int preInputRawRolls,
            int connectionRolls,
            int secBeforeWipRolls,
            int secCompleteRolls,
            double ruleWipMetric,
            double ruleWipThreshold,
            int rollIndex,
            int rollTotal) {
        preInputRaw.setCount(preInputRawRolls, rollTotal, "", false);
        connection.setCount(connectionRolls, rollTotal, "", false);
        String wipSubtitle =
                ruleWipMetric >= 0 && ruleWipThreshold > 0
                        ? String.format("L13判定WIP %.0f/%.0f", ruleWipMetric, ruleWipThreshold)
                        : "";
        boolean wipOver = ruleWipMetric >= ruleWipThreshold && ruleWipThreshold > 0;
        secBeforeWip.setCount(secBeforeWipRolls, rollTotal, wipSubtitle, wipOver);
        secComplete.setCount(secCompleteRolls, rollTotal, "", false);
        if (rollIndex > 0 && rollTotal > 0) {
            preInputRaw.setSubtitle("試走ロール " + rollIndex + "/" + rollTotal);
        } else {
            preInputRaw.setSubtitle("");
        }
    }

    private static final class StageColumn {
        private final Label title = new Label();
        private final Label countLabel = new Label("0");
        private final Label subtitle = new Label();
        private final FlowPane icons = new FlowPane(3, 3);
        private final VBox root = new VBox(4);
        private final Color fill;

        StageColumn(String titleText, Color fill) {
            this.fill = fill;
            title.setText(titleText);
            title.setStyle("-fx-font-weight: bold; -fx-font-size: 11px;");
            countLabel.setStyle("-fx-font-size: 12px;");
            subtitle.setStyle("-fx-text-fill: #666; -fx-font-size: 10px;");
            icons.setPrefWrapLength(140);
            icons.setAlignment(Pos.CENTER_LEFT);
            root.getChildren().addAll(title, countLabel, subtitle, icons);
        }

        void setSubtitle(String text) {
            subtitle.setText(text != null ? text : "");
        }

        void setCount(int count, int rollTotal, String extraSubtitle, boolean highlightOver) {
            if (extraSubtitle != null && !extraSubtitle.isBlank()) {
                subtitle.setText(extraSubtitle);
            }
            if (highlightOver) {
                countLabel.setStyle("-fx-font-size: 12px; -fx-text-fill: #E74C3C; -fx-font-weight: bold;");
            } else {
                countLabel.setStyle("-fx-font-size: 12px;");
            }
            int n = Math.max(0, count);
            int cap = rollTotal > 0 ? rollTotal : MAX_ICONS;
            countLabel.setText(String.valueOf(n));
            icons.getChildren().clear();
            int show = Math.min(n, Math.min(cap, MAX_ICONS));
            Color iconFill = highlightOver ? Color.web("#E74C3C") : fill;
            for (int i = 0; i < show; i++) {
                Circle c = new Circle(ICON_RADIUS, iconFill);
                icons.getChildren().add(c);
            }
            if (n > show) {
                icons.getChildren().add(new Label("+" + (n - show)));
            }
        }
    }
}
