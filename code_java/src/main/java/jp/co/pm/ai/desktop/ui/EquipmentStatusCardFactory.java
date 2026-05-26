package jp.co.pm.ai.desktop.ui;

import java.util.List;
import java.util.function.Function;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Label;
import javafx.scene.control.OverrunStyle;
import javafx.scene.control.Tooltip;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;

import jp.co.pm.ai.desktop.config.EquipmentStatusDashboardAppearancePrefs;
import jp.co.pm.ai.desktop.config.PersonBadgeStyle;
import jp.co.pm.ai.desktop.io.actuals.EquipmentMachineStatus;
import jp.co.pm.ai.desktop.io.gantt.PersonNameHeuristics;
import jp.co.pm.ai.desktop.io.gantt.PersonNameBadgeText;

/** 設備現状ダッシュボードの機械カード Node 生成。 */
public final class EquipmentStatusCardFactory {

    public record DisplayOptions(
            boolean showAladdinPlans,
            boolean showDispatchPlans,
            String actualDateLabel,
            String planDateLabel,
            /** 実績日が当日のときのみ {@code true}。「稼働中」チップ表示に使う。 */
            boolean showRunningStatusChip) {

        public static DisplayOptions defaults() {
            return new DisplayOptions(true, true, "", "", true);
        }
    }

    private EquipmentStatusCardFactory() {}

    public static VBox createCard(
            EquipmentMachineStatus status,
            DisplayOptions opts,
            EquipmentStatusDashboardAppearancePrefs appearance,
            Function<String, PersonBadgeStyle> badgeStyleResolver,
            boolean fullscreen) {
        EquipmentStatusDashboardAppearancePrefs ap =
                appearance != null
                        ? appearance
                        : EquipmentStatusDashboardAppearancePrefs.defaults();
        double width = EquipmentStatusDashboardAppearanceApplier.snappedCardWidth(ap, fullscreen);
        VBox card = new VBox(8.0);
        card.getStyleClass().add("pm-equipment-status-card");
        EquipmentStatusDashboardAppearanceApplier.applyCardShell(card, ap, fullscreen);

        HBox header = new HBox(8.0);
        header.setAlignment(Pos.CENTER_LEFT);
        Label machine = new Label(status.machineName());
        machine.getStyleClass().add("pm-equipment-status-machine");
        EquipmentStatusDashboardAppearanceApplier.applyMachineLabel(machine, ap);
        machine.setMaxWidth(shouldShowStatusChip(status, opts) ? width - 100 : width - ap.cardPadding() * 2);
        machine.setTextOverrun(OverrunStyle.ELLIPSIS);
        header.getChildren().add(machine);
        if (shouldShowStatusChip(status, opts)) {
            Region spacer = new Region();
            HBox.setHgrow(spacer, Priority.ALWAYS);
            Label chip = new Label(statusLabel(status.status()));
            chip.getStyleClass().addAll("pm-equipment-status-chip", chipStyleClass(status.status()));
            EquipmentStatusDashboardAppearanceApplier.applyLabelFont(chip, ap, ap.planFontPx());
            Tooltip.install(chip, new Tooltip(statusChipTooltip(status.status())));
            header.getChildren().addAll(spacer, chip);
        }

        card.getChildren().add(header);

        if (status.status() == EquipmentMachineStatus.Status.STOPPED) {
            appendSectionRule(card, width, ap);
            StackPane stoppedPane = new StackPane();
            stoppedPane.getStyleClass().add("pm-equipment-status-stopped-pane");
            stoppedPane.setMinHeight(Math.max(80, ap.chartSizePx() * 1.1));
            Label stopped = new Label("停機");
            stopped.getStyleClass().add("pm-equipment-status-stopped-label");
            EquipmentStatusDashboardAppearanceApplier.applyLabelFont(
                    stopped, ap, ap.machineFontPx() * 1.75);
            stoppedPane.getChildren().add(stopped);
            card.getChildren().add(stoppedPane);
        } else {
            status.actualTask()
                    .ifPresent(
                            task -> {
                                appendSectionRule(card, width, ap);
                                card.getChildren()
                                        .add(
                                                buildLatestActualSection(
                                                        task,
                                                        width,
                                                        ap,
                                                        badgeStyleResolver));
                                appendSectionRule(card, width, ap);
                                card.getChildren()
                                        .add(
                                                buildDayAladdinAchievementSection(
                                                        task, opts, width, ap));
                            });
        }

        appendPlanSection(
                card,
                "アラジン予定",
                opts != null && opts.showAladdinPlans(),
                status.aladdinPlans(),
                opts != null ? opts.planDateLabel() : "",
                width,
                ap);
        appendPlanSection(
                card,
                "配台予定",
                opts != null && opts.showDispatchPlans(),
                status.dispatchPlans(),
                opts != null ? opts.planDateLabel() : "",
                width,
                ap);

        return card;
    }

    private static void appendSectionRule(
            VBox card, double cardWidth, EquipmentStatusDashboardAppearancePrefs ap) {
        double contentWidth = cardWidth - ap.cardPadding() * 2;
        Region rule = new Region();
        rule.getStyleClass().add("pm-equipment-status-section-rule");
        rule.setMinWidth(contentWidth);
        rule.setPrefWidth(contentWidth);
        rule.setMaxWidth(contentWidth);
        rule.setMinHeight(1);
        rule.setPrefHeight(1);
        rule.setMaxHeight(1);
        VBox.setMargin(rule, new Insets(4, 0, 4, 0));
        card.getChildren().add(rule);
    }

    private static VBox buildLatestActualSection(
            EquipmentMachineStatus.ActualTaskRow task,
            double cardWidth,
            EquipmentStatusDashboardAppearancePrefs ap,
            Function<String, PersonBadgeStyle> badgeStyleResolver) {
        double contentWidth = cardWidth - ap.cardPadding() * 2;
        VBox section = new VBox(2.0);
        Label head = new Label("直近実績");
        head.getStyleClass().add("pm-equipment-status-plan-head");
        EquipmentStatusDashboardAppearanceApplier.applyLabelFont(head, ap, ap.planFontPx());
        Label meta = new Label(metaLine(task));
        meta.getStyleClass().add("pm-equipment-status-meta");
        EquipmentStatusDashboardAppearanceApplier.applyLabelFont(meta, ap, ap.metaFontPx());
        meta.setWrapText(true);
        meta.setMaxWidth(contentWidth);
        section.getChildren().addAll(head, meta);
        String badgeText = PersonNameBadgeText.badgeTwoFromRawName(task.memberRaw());
        if (!badgeText.isBlank()
                && PersonNameHeuristics.looksLikePersonName(task.memberRaw())) {
            PersonBadgeStyle st =
                    badgeStyleResolver != null
                            ? badgeStyleResolver.apply(
                                    PersonNameBadgeText.surnameLabelOnly(task.memberRaw()))
                            : PersonBadgeStyle.defaultStyle();
            section.getChildren()
                    .add(
                            PersonBadgeNodeFactory.createBadge(
                                    badgeText, st, 1.0, Math.max(10, ap.metaFontPx())));
        }
        return section;
    }

    private static VBox buildDayAladdinAchievementSection(
            EquipmentMachineStatus.ActualTaskRow task,
            DisplayOptions opts,
            double cardWidth,
            EquipmentStatusDashboardAppearancePrefs ap) {
        double contentWidth = cardWidth - ap.cardPadding() * 2;
        VBox section = new VBox(4.0);
        section.setAlignment(Pos.CENTER_LEFT);

        Label head = new Label(achievementSectionTitle(opts));
        head.getStyleClass().add("pm-equipment-status-plan-head");
        head.setWrapText(true);
        head.setMaxWidth(contentWidth);
        EquipmentStatusDashboardAppearanceApplier.applyLabelFont(head, ap, ap.planFontPx());

        StackPane chartPane =
                EquipmentStatusDashboardAppearanceApplier.buildPieChart(task.completionPct(), ap);
        Tooltip.install(
                chartPane,
                new Tooltip(
                        "当機械の当日実績合計 ÷ 当日アラジン計画合計。"
                                + " 直近実績の依頼NO単位の進捗ではありません。"));

        Label foot = new Label("当機械 · 実績合計 ÷ 当日アラジン計画合計");
        foot.getStyleClass().add("pm-equipment-status-achievement-foot");
        foot.setWrapText(true);
        foot.setMaxWidth(contentWidth);
        EquipmentStatusDashboardAppearanceApplier.applyLabelFont(
                foot, ap, Math.max(9, ap.planFontPx() - 1));

        section.getChildren().add(head);
        section.getChildren().add(chartPane);
        section.getChildren().add(foot);
        return section;
    }

    private static String achievementSectionTitle(DisplayOptions opts) {
        String date = opts != null ? nz(opts.actualDateLabel()) : "—";
        if ("—".equals(date)) {
            return "当日アラジン達成率（機械合計）";
        }
        return "当日アラジン達成率（" + date + "・機械合計）";
    }

    private static String statusChipTooltip(EquipmentMachineStatus.Status status) {
        return switch (status) {
            case STOPPED -> "選択した実績日に、この機械の加工実績がありません。";
            case RUNNING ->
                    "当日アラジン計画（当機械合計）に対し、実績合計が未達です。依頼NO単位の完了ではありません。";
            case COMPLETED ->
                    "当日アラジン計画（当機械合計）を実績合計が上回った状態です。依頼NO単位の完了ではありません。";
        };
    }

    private static void appendPlanSection(
            VBox card,
            String title,
            boolean visible,
            List<EquipmentMachineStatus.PlanLine> lines,
            String dateLabel,
            double width,
            EquipmentStatusDashboardAppearancePrefs ap) {
        if (!visible) {
            return;
        }
        appendSectionRule(card, width, ap);
        Label head = new Label(title + (dateLabel.isBlank() ? "" : " (" + dateLabel + ")"));
        head.getStyleClass().add("pm-equipment-status-plan-head");
        EquipmentStatusDashboardAppearanceApplier.applyLabelFont(head, ap, ap.planFontPx());
        card.getChildren().add(head);
        if (lines == null || lines.isEmpty()) {
            Label empty = new Label("—");
            empty.getStyleClass().add("pm-equipment-status-plan-line");
            EquipmentStatusDashboardAppearanceApplier.applyLabelFont(empty, ap, ap.planFontPx());
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
            EquipmentStatusDashboardAppearanceApplier.applyLabelFont(row, ap, ap.planFontPx());
            row.setWrapText(true);
            row.setMaxWidth(width - ap.cardPadding() * 2);
            card.getChildren().add(row);
        }
    }

    private static boolean shouldShowStatusChip(
            EquipmentMachineStatus status, DisplayOptions opts) {
        if (status.status() == EquipmentMachineStatus.Status.RUNNING) {
            return opts == null || opts.showRunningStatusChip();
        }
        return true;
    }

    private static String statusLabel(EquipmentMachineStatus.Status status) {
        return switch (status) {
            case STOPPED -> "停機";
            case RUNNING -> "稼働中";
            case COMPLETED -> "予定達成";
        };
    }

    private static String chipStyleClass(EquipmentMachineStatus.Status status) {
        return switch (status) {
            case STOPPED -> "pm-equipment-status-chip-stopped";
            case RUNNING -> "pm-equipment-status-chip-running";
            case COMPLETED -> "pm-equipment-status-chip-completed";
        };
    }

    /**
     * 選択日に実績・予定が1件も無いときの説明パネル（工場非稼働日など）。
     *
     * @param sourcesLoaded {@code true} なら読込済みだが該当0件、{@code false} なら未読込
     */
    public static VBox createEmptyState(
            String actualDateLabel, String planDateLabel, boolean sourcesLoaded, boolean fullscreen) {
        return createEmptyState(actualDateLabel, planDateLabel, sourcesLoaded, fullscreen, "");
    }

    /**
     * @param loadErrorDetail 空でなければ読込失敗パネルとして表示
     */
    public static VBox createEmptyState(
            String actualDateLabel,
            String planDateLabel,
            boolean sourcesLoaded,
            boolean fullscreen,
            String loadErrorDetail) {
        VBox box = new VBox(10.0);
        box.setAlignment(Pos.CENTER);
        box.setPadding(new Insets(32, 24, 32, 24));
        box.setMaxWidth(fullscreen ? 640 : 560);
        box.getStyleClass().add("pm-equipment-status-empty");

        Label title = new Label("表示する機械がありません");
        title.getStyleClass().add("pm-equipment-status-empty-title");

        String detailText;
        String hintText;
        if (loadErrorDetail != null && !loadErrorDetail.isBlank()) {
            title.setText("ダッシュボード読込エラー");
            detailText = loadErrorDetail.strip();
            hintText =
                    "［再読込］を押すか、「実行・ログ」タブで詳細（スタックトレース）を確認してください。";
        } else if (sourcesLoaded) {
            detailText =
                    "実績 "
                            + nz(actualDateLabel)
                            + "・予定 "
                            + nz(planDateLabel)
                            + " のいずれかにデータがある機械が0件です。";
            hintText =
                    "工場が休業・非稼働の日は、加工実績もアラジン／配台予定も無いことがあります。"
                            + "［当日］ボタンやカレンダーで日付を変えて確認してください。";
        } else {
            detailText = "ソースを読み込んでいません。";
            hintText = "［再読込］を押すか、タブを開き直してください。";
        }

        Label detail = new Label(detailText);
        detail.getStyleClass().add("pm-equipment-status-empty-detail");
        detail.setWrapText(true);
        detail.setMaxWidth(fullscreen ? 600 : 520);
        detail.setAlignment(Pos.CENTER);

        Label hint = new Label(hintText);
        hint.getStyleClass().add("pm-equipment-status-empty-hint");
        hint.setWrapText(true);
        hint.setMaxWidth(fullscreen ? 600 : 520);
        hint.setAlignment(Pos.CENTER);

        box.getChildren().addAll(title, detail, hint);
        return box;
    }

    private static String metaLine(EquipmentMachineStatus.ActualTaskRow task) {
        String line = task.requestNo() + " · " + task.processName();
        if (task.memberRaw() != null
                && PersonNameHeuristics.looksLikePersonName(task.memberRaw())) {
            line += " · " + task.memberRaw().strip();
        }
        return line;
    }

    private static String nz(String s) {
        return s != null && !s.isBlank() ? s.strip() : "—";
    }
}
