package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.EnumMap;
import java.util.List;
import java.util.Map;
import java.util.function.Function;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
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

    /** 予定セクションに並べる最大行数。超過分は「他 n 件」に畳む。 */
    static final int MAX_PLAN_LINES = 5;

    private static final String ACHIEVEMENT_HELP =
            "当機械の当日実績合計 ÷ 当日アラジン計画合計。依頼NO単位の進捗ではありません。";

    /** ステータス3種と達成率の説明は内容が固定なので、カードごとに作らず使い回す。 */
    private static final Map<EquipmentMachineStatus.Status, Tooltip> CHIP_TOOLTIPS =
            new EnumMap<>(EquipmentMachineStatus.Status.class);

    private static Tooltip achievementTooltip;

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
        double contentWidth = Math.max(40, width - ap.cardPadding() * 2);
        VBox card = new VBox(8.0);
        card.getStyleClass().addAll("pm-equipment-status-card", cardStateStyleClass(status.status()));
        card.setFocusTraversable(true);
        card.setAccessibleText(cardAccessibleText(status));
        card.setAccessibleHelp(ACHIEVEMENT_HELP);
        EquipmentStatusDashboardAppearanceApplier.applyCardShell(card, ap, fullscreen);

        card.getChildren().add(buildHeader(status, opts, ap));

        if (status.status() == EquipmentMachineStatus.Status.STOPPED) {
            StackPane stoppedPane = new StackPane();
            stoppedPane.getStyleClass().add("pm-equipment-status-stopped-pane");
            stoppedPane.setMinHeight(Math.max(80, ap.chartSizePx() * 1.1));
            Label stopped = new Label("停機");
            stopped.getStyleClass().add("pm-equipment-status-stopped-label");
            EquipmentStatusDashboardAppearanceApplier.applyLabelFont(
                    stopped, ap, ap.machineFontPx() * 1.75, true);
            stoppedPane.getChildren().add(stopped);
            card.getChildren().add(stoppedPane);
        } else {
            status.actualTask()
                    .ifPresent(
                            task -> {
                                card.getChildren()
                                        .add(
                                                buildLatestActualSection(
                                                        task,
                                                        contentWidth,
                                                        ap,
                                                        badgeStyleResolver));
                                card.getChildren()
                                        .add(buildAchievementSection(task, contentWidth, ap));
                            });
        }

        boolean showAladdin = opts != null && opts.showAladdinPlans();
        boolean showDispatch = opts != null && opts.showDispatchPlans();
        if (showAladdin || showDispatch) {
            appendSectionRule(card);
        }
        appendPlanSection(card, "アラジン予定", showAladdin, status.aladdinPlans(), contentWidth, ap);
        appendPlanSection(card, "配台予定", showDispatch, status.dispatchPlans(), contentWidth, ap);

        return card;
    }

    private static HBox buildHeader(
            EquipmentMachineStatus status,
            DisplayOptions opts,
            EquipmentStatusDashboardAppearancePrefs ap) {
        HBox header = new HBox(8.0);
        header.setAlignment(Pos.CENTER_LEFT);
        Label machine = new Label(status.machineName());
        machine.getStyleClass().add("pm-equipment-status-machine");
        EquipmentStatusDashboardAppearanceApplier.applyMachineLabel(machine, ap);
        machine.setTextOverrun(OverrunStyle.ELLIPSIS);
        machine.setMinWidth(0);
        machine.setMaxWidth(Double.MAX_VALUE);
        HBox.setHgrow(machine, Priority.ALWAYS);
        header.getChildren().add(machine);
        if (shouldShowStatusChip(status, opts)) {
            Label chip = new Label(statusLabel(status.status()));
            chip.getStyleClass().addAll("pm-equipment-status-chip", chipStyleClass(status.status()));
            EquipmentStatusDashboardAppearanceApplier.applyLabelFont(
                    chip, ap, Math.max(11, ap.planFontPx() + 1), true);
            chip.setMinWidth(Region.USE_PREF_SIZE);
            Tooltip.install(chip, chipTooltip(status.status()));
            header.getChildren().add(chip);
        }
        return header;
    }

    /** カード内ブロックの境界線。幅は VBox の fillWidth に任せる。 */
    private static void appendSectionRule(VBox card) {
        Region rule = new Region();
        rule.getStyleClass().add("pm-equipment-status-section-rule");
        rule.setMaxWidth(Double.MAX_VALUE);
        VBox.setMargin(rule, new Insets(2, 0, 2, 0));
        card.getChildren().add(rule);
    }

    private static VBox buildLatestActualSection(
            EquipmentMachineStatus.ActualTaskRow task,
            double contentWidth,
            EquipmentStatusDashboardAppearancePrefs ap,
            Function<String, PersonBadgeStyle> badgeStyleResolver) {
        VBox section = new VBox(2.0);
        Label head = new Label("直近実績");
        head.getStyleClass().add("pm-equipment-status-plan-head");
        EquipmentStatusDashboardAppearanceApplier.applyLabelFont(head, ap, ap.planFontPx(), true);
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

    private static VBox buildAchievementSection(
            EquipmentMachineStatus.ActualTaskRow task,
            double contentWidth,
            EquipmentStatusDashboardAppearancePrefs ap) {
        VBox section = new VBox(4.0);
        section.setAlignment(Pos.TOP_LEFT);

        Label head = new Label("アラジン達成率");
        head.getStyleClass().add("pm-equipment-status-plan-head");
        head.setMaxWidth(contentWidth);
        EquipmentStatusDashboardAppearanceApplier.applyLabelFont(head, ap, ap.planFontPx(), true);

        EquipmentStatusProgressRing ring = new EquipmentStatusProgressRing();
        ring.update(task.completionPct(), ap);
        Tooltip.install(ring, achievementTooltip());

        HBox ringRow = new HBox(ring);
        ringRow.setAlignment(Pos.CENTER);
        ringRow.setMaxWidth(Double.MAX_VALUE);

        section.getChildren().addAll(head, ringRow);
        return section;
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

    private static Tooltip chipTooltip(EquipmentMachineStatus.Status status) {
        return CHIP_TOOLTIPS.computeIfAbsent(status, k -> new Tooltip(statusChipTooltip(k)));
    }

    private static Tooltip achievementTooltip() {
        if (achievementTooltip == null) {
            achievementTooltip = new Tooltip(ACHIEVEMENT_HELP);
        }
        return achievementTooltip;
    }

    private static void appendPlanSection(
            VBox card,
            String title,
            boolean visible,
            List<EquipmentMachineStatus.PlanLine> lines,
            double contentWidth,
            EquipmentStatusDashboardAppearancePrefs ap) {
        if (!visible) {
            return;
        }
        Label head = new Label(title);
        head.getStyleClass().add("pm-equipment-status-plan-head");
        EquipmentStatusDashboardAppearanceApplier.applyLabelFont(head, ap, ap.planFontPx(), true);
        card.getChildren().add(head);
        if (lines == null || lines.isEmpty()) {
            Label empty = new Label("—");
            empty.getStyleClass().add("pm-equipment-status-plan-line");
            EquipmentStatusDashboardAppearanceApplier.applyLabelFont(empty, ap, ap.planFontPx());
            card.getChildren().add(empty);
            return;
        }
        int shown = Math.min(lines.size(), MAX_PLAN_LINES);
        for (int i = 0; i < shown; i++) {
            Label row = new Label(planLineText(lines.get(i)));
            row.getStyleClass().add("pm-equipment-status-plan-line");
            EquipmentStatusDashboardAppearanceApplier.applyLabelFont(row, ap, ap.planFontPx());
            row.setWrapText(true);
            row.setMaxWidth(contentWidth);
            card.getChildren().add(row);
        }
        if (lines.size() > shown) {
            Label more = new Label("他 " + (lines.size() - shown) + " 件");
            more.getStyleClass().addAll("pm-equipment-status-plan-line", "pm-equipment-status-plan-more");
            EquipmentStatusDashboardAppearanceApplier.applyLabelFont(more, ap, ap.planFontPx());
            Tooltip rest = new Tooltip(remainingPlanLinesText(lines, shown));
            rest.setWrapText(true);
            rest.setMaxWidth(480);
            Tooltip.install(more, rest);
            card.getChildren().add(more);
        }
    }

    static String planLineText(EquipmentMachineStatus.PlanLine line) {
        return line.requestNo() + " · " + line.processName() + " · " + line.qtyM() + "m";
    }

    static String remainingPlanLinesText(
            List<EquipmentMachineStatus.PlanLine> lines, int alreadyShown) {
        List<String> rest = new ArrayList<>();
        for (int i = alreadyShown; i < lines.size(); i++) {
            rest.add(planLineText(lines.get(i)));
        }
        return String.join("\n", rest);
    }

    static boolean shouldShowStatusChip(EquipmentMachineStatus status, DisplayOptions opts) {
        if (status.status() == EquipmentMachineStatus.Status.RUNNING) {
            return opts == null || opts.showRunningStatusChip();
        }
        return true;
    }

    static String statusLabel(EquipmentMachineStatus.Status status) {
        return switch (status) {
            case STOPPED -> "停機";
            case RUNNING -> "稼働中";
            case COMPLETED -> "予定達成";
        };
    }

    static String chipStyleClass(EquipmentMachineStatus.Status status) {
        return switch (status) {
            case STOPPED -> "pm-equipment-status-chip-stopped";
            case RUNNING -> "pm-equipment-status-chip-running";
            case COMPLETED -> "pm-equipment-status-chip-completed";
        };
    }

    /** カード左端のアクセントバー色を状態別に切り替えるためのクラス。 */
    static String cardStateStyleClass(EquipmentMachineStatus.Status status) {
        if (status == null) {
            return "pm-equipment-status-card-running";
        }
        return switch (status) {
            case STOPPED -> "pm-equipment-status-card-stopped";
            case RUNNING -> "pm-equipment-status-card-running";
            case COMPLETED -> "pm-equipment-status-card-completed";
        };
    }

    static String cardAccessibleText(EquipmentMachineStatus status) {
        StringBuilder sb = new StringBuilder();
        sb.append(status.machineName() != null ? status.machineName() : "機械名不明");
        sb.append("、").append(statusLabel(status.status()));
        status.actualTask()
                .ifPresent(
                        task ->
                                sb.append("、アラジン達成率 ")
                                        .append(Math.round(Math.max(0, task.completionPct())))
                                        .append("パーセント"));
        return sb.toString();
    }

    /**
     * 選択日に実績・予定が1件も無いとき、または読込に失敗したときの説明パネル。
     *
     * @param sourcesLoaded {@code true} なら読込済みだが該当0件、{@code false} なら未読込
     * @param loadErrorDetail 空でなければ読込失敗パネルとして表示
     * @param onReload {@code null} でなければ「再読込」ボタンを置く
     */
    public static VBox createEmptyState(
            String actualDateLabel,
            String planDateLabel,
            boolean sourcesLoaded,
            boolean fullscreen,
            String loadErrorDetail,
            Runnable onReload) {
        boolean isError = loadErrorDetail != null && !loadErrorDetail.isBlank();
        VBox box = new VBox(10.0);
        box.setAlignment(Pos.CENTER);
        box.setPadding(new Insets(32, 24, 32, 24));
        box.setMaxWidth(fullscreen ? 640 : 560);
        box.getStyleClass().add("pm-equipment-status-empty");
        if (isError) {
            box.getStyleClass().add("pm-equipment-status-empty-error");
        }

        Label title = new Label(isError ? "ダッシュボード読込エラー" : "表示する機械がありません");
        title.getStyleClass().add("pm-equipment-status-empty-title");
        if (isError) {
            title.getStyleClass().add("pm-equipment-status-error");
        }

        String detailText;
        String hintText;
        if (isError) {
            detailText = loadErrorDetail.strip();
            hintText = "「実行・ログ」タブで詳細（スタックトレース）を確認できます。";
        } else if (sourcesLoaded) {
            detailText =
                    "実績 "
                            + nz(actualDateLabel)
                            + "・予定 "
                            + nz(planDateLabel)
                            + " のいずれかにデータがある機械が0件です。";
            hintText =
                    "工場が休業・非稼働の日は、加工実績もアラジン／配台予定も無いことがあります。"
                            + "絞込を解除するか、日付を変えて確認してください。";
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
        if (onReload != null) {
            Button reload = new Button("再読込");
            reload.setDefaultButton(false);
            reload.setOnAction(e -> onReload.run());
            box.getChildren().add(reload);
        }
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

    static String nz(String s) {
        return s != null && !s.isBlank() ? s.strip() : "—";
    }
}
