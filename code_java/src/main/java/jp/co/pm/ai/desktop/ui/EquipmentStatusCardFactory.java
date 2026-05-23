package jp.co.pm.ai.desktop.ui;

import java.util.Locale;
import java.util.List;
import java.util.function.Function;

import javafx.collections.FXCollections;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.chart.PieChart;
import javafx.scene.control.Label;
import javafx.scene.control.OverrunStyle;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;

import jp.co.pm.ai.desktop.config.PersonBadgeStyle;
import jp.co.pm.ai.desktop.io.actuals.EquipmentMachineStatus;
import jp.co.pm.ai.desktop.io.gantt.PersonNameBadgeText;

/** 設備現状ダッシュボードの機械カード Node 生成。 */
public final class EquipmentStatusCardFactory {

    public record DisplayOptions(
            boolean showAladdinPlans,
            boolean showDispatchPlans,
            String actualDateLabel,
            String planDateLabel) {

        public static DisplayOptions defaults() {
            return new DisplayOptions(true, true, "", "");
        }
    }

    private static final double CARD_WIDTH = 280.0;
    private static final double CARD_WIDTH_FULLSCREEN = 340.0;
    private static final double CHART_SIZE = 96.0;

    private EquipmentStatusCardFactory() {}

    public static VBox createCard(
            EquipmentMachineStatus status,
            DisplayOptions opts,
            Function<String, PersonBadgeStyle> badgeStyleResolver,
            boolean fullscreen) {
        double width = fullscreen ? CARD_WIDTH_FULLSCREEN : CARD_WIDTH;
        VBox card = new VBox(8.0);
        card.getStyleClass().add("pm-equipment-status-card");
        card.setPrefWidth(width);
        card.setMinWidth(width);
        card.setMaxWidth(width);
        card.setPadding(new Insets(12.0));

        HBox header = new HBox(8.0);
        header.setAlignment(Pos.CENTER_LEFT);
        Label machine = new Label(status.machineName());
        machine.getStyleClass().add("pm-equipment-status-machine");
        machine.setMaxWidth(width - 100);
        machine.setTextOverrun(OverrunStyle.ELLIPSIS);
        Region spacer = new Region();
        HBox.setHgrow(spacer, Priority.ALWAYS);
        Label chip = new Label(statusLabel(status.status()));
        chip.getStyleClass().addAll("pm-equipment-status-chip", chipStyleClass(status.status()));
        header.getChildren().addAll(machine, spacer, chip);

        card.getChildren().add(header);

        if (status.status() == EquipmentMachineStatus.Status.STOPPED) {
            StackPane stoppedPane = new StackPane();
            stoppedPane.getStyleClass().add("pm-equipment-status-stopped-pane");
            stoppedPane.setMinHeight(120);
            Label stopped = new Label("停機");
            stopped.getStyleClass().add("pm-equipment-status-stopped-label");
            stoppedPane.getChildren().add(stopped);
            card.getChildren().add(stoppedPane);
        } else {
            status.actualTask()
                    .ifPresent(
                            task -> {
                                Label meta =
                                        new Label(
                                                task.requestNo()
                                                        + " · "
                                                        + task.processName());
                                meta.getStyleClass().add("pm-equipment-status-meta");
                                meta.setWrapText(true);
                                meta.setMaxWidth(width - 24);
                                card.getChildren().add(meta);

                                StackPane chartPane = buildPieChart(task.completionPct());
                                card.getChildren().add(chartPane);

                                String badgeText =
                                        PersonNameBadgeText.badgeTwoFromRawName(task.memberRaw());
                                if (!badgeText.isBlank()) {
                                    PersonBadgeStyle st =
                                            badgeStyleResolver != null
                                                    ? badgeStyleResolver.apply(
                                                            PersonNameBadgeText.surnameLabelOnly(
                                                                    task.memberRaw()))
                                                    : PersonBadgeStyle.defaultStyle();
                                    card.getChildren()
                                            .add(
                                                    PersonBadgeNodeFactory.createBadge(
                                                            badgeText, st, 1.0, 12.0));
                                }
                            });
        }

        appendPlanSection(
                card,
                "アラジン予定",
                opts != null && opts.showAladdinPlans(),
                status.aladdinPlans(),
                opts != null ? opts.planDateLabel() : "",
                width);
        appendPlanSection(
                card,
                "配台予定",
                opts != null && opts.showDispatchPlans(),
                status.dispatchPlans(),
                opts != null ? opts.planDateLabel() : "",
                width);

        if (opts != null && !opts.actualDateLabel().isBlank()) {
            Label actualLbl = new Label("実績: " + opts.actualDateLabel());
            actualLbl.getStyleClass().add("pm-equipment-status-date-foot");
            card.getChildren().add(actualLbl);
        }

        return card;
    }

    private static void appendPlanSection(
            VBox card,
            String title,
            boolean visible,
            List<EquipmentMachineStatus.PlanLine> lines,
            String dateLabel,
            double width) {
        if (!visible) {
            return;
        }
        Label head = new Label(title + (dateLabel.isBlank() ? "" : " (" + dateLabel + ")"));
        head.getStyleClass().add("pm-equipment-status-plan-head");
        card.getChildren().add(head);
        if (lines == null || lines.isEmpty()) {
            Label empty = new Label("—");
            empty.getStyleClass().add("pm-equipment-status-plan-line");
            card.getChildren().add(empty);
            return;
        }
        for (EquipmentMachineStatus.PlanLine line : lines) {
            Label row =
                    new Label(
                            line.requestNo()
                                    + " · "
                                    + line.processName()
                                    + " · "
                                    + line.qtyM()
                                    + "m");
            row.getStyleClass().add("pm-equipment-status-plan-line");
            row.setWrapText(true);
            row.setMaxWidth(width - 24);
            card.getChildren().add(row);
        }
    }

    private static StackPane buildPieChart(double completionPct) {
        double done = Math.max(0.0, Math.min(100.0, completionPct));
        double remain = 100.0 - done;
        PieChart chart =
                new PieChart(
                        FXCollections.observableArrayList(
                                new PieChart.Data("完了", done),
                                new PieChart.Data("残り", remain > 0 ? remain : 0.01)));
        chart.setAnimated(false);
        chart.setLegendVisible(false);
        chart.setLabelsVisible(false);
        chart.setPrefSize(CHART_SIZE, CHART_SIZE);
        chart.setMinSize(CHART_SIZE, CHART_SIZE);
        chart.setMaxSize(CHART_SIZE, CHART_SIZE);
        chart.getStyleClass().add("pm-equipment-status-pie");

        Label pct = new Label(String.format("%.0f%%", done));
        pct.getStyleClass().add("pm-equipment-status-pct-label");
        StackPane pane = new StackPane(chart, pct);
        pane.setAlignment(Pos.CENTER);
        return pane;
    }

    private static String statusLabel(EquipmentMachineStatus.Status status) {
        return switch (status) {
            case STOPPED -> "停機";
            case RUNNING -> "稼働中";
            case COMPLETED -> "完了";
        };
    }

    private static String chipStyleClass(EquipmentMachineStatus.Status status) {
        return switch (status) {
            case STOPPED -> "pm-equipment-status-chip-stopped";
            case RUNNING -> "pm-equipment-status-chip-running";
            case COMPLETED -> "pm-equipment-status-chip-completed";
        };
    }
}
