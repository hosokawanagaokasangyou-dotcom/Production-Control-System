package jp.co.pm.ai.desktop.config;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;

import jp.co.pm.ai.desktop.MainShellTabId;

/**
 * Default grouped layout for main-shell tabs (tab organizer baseline when session has no layout).
 *
 * <p>Tab-management initial ordering / grouping for documentation is in
 * {@code .cursor/rules/main-shell-tab-management.mdc}; it is not duplicated here as a normative catalog.
 *
 * <p>Add new {@link MainShellTabId} keys at the end of {@link #DEFAULT_FLAT_TAB_KEY_ORDER} (before tab organizer).
 *
 * <p>Rulebook: {@code .cursor/rules/main-shell-tab-management.mdc}
 *
 * <p>並び・グループ・色は {@code init_setting/session_defaults_*.json} のグローバル既定と同期する。
 */
public final class MainShellTabLayoutDefaults {

    private MainShellTabLayoutDefaults() {}

    /** Flat tab key order (reset-flat button and merge order for missing keys). */
    public static final List<String> DEFAULT_FLAT_TAB_KEY_ORDER =
            List.of(
                    MainShellTabId.EQUIPMENT_STATUS_DASHBOARD.key(),
                    MainShellTabId.REMOTE_DESKTOP.key(),
                    MainShellTabId.REQUEST_FORM_INPUT.key(),
                    MainShellTabId.REQUEST_FORM_PIPELINE_CHECK.key(),
                    MainShellTabId.COMPANY_CALENDAR.key(),
                    MainShellTabId.MEMBER_ATTENDANCE.key(),
                    MainShellTabId.MACHINE_CALENDAR.key(),
                    MainShellTabId.MASTER_DISPATCH_SHEETS.key(),
                    MainShellTabId.RUN.key(),
                    MainShellTabId.PLAN_INPUT.key(),
                    MainShellTabId.EQUIPMENT_GANTT_GRAPHIC.key(),
                    MainShellTabId.DELIVERY_CALENDAR_VIEW.key(),
                    MainShellTabId.OPERATOR_CARD.key(),
                    MainShellTabId.CODE_LOOKUP_TABLES.key(),
                    MainShellTabId.UI_BADGE_DESIGN.key(),
                    MainShellTabId.PUSH_BUTTON_DESIGN.key(),
                    MainShellTabId.GANTT_PERSON_BADGE_DESIGN.key(),
                    MainShellTabId.REQUEST_FORM_PREVIEW_BADGE_DESIGN.key(),
                    MainShellTabId.PLAN_RESULT_VIEWER.key(),
                    MainShellTabId.STAGE1_PREVIEW.key(),
                    MainShellTabId.RESULT_DISPATCH.key(),
                    MainShellTabId.MASTER_SUMMARY.key(),
                    MainShellTabId.EXCLUDE_RULES.key(),
                    MainShellTabId.SPECIAL_RULES.key(),
                    MainShellTabId.PLAN_WORKSPACE_HISTORY.key(),
                    MainShellTabId.MEMORY_SETTINGS.key(),
                    MainShellTabId.API_MODEL_BENCHMARK.key(),
                    MainShellTabId.ACTUALS_STATUS.key(),
                    MainShellTabId.DAILY_REPORT_CSV_VIEW.key(),
                    MainShellTabId.PIPELINE_EXECUTION_TIMING.key(),
                    MainShellTabId.CACHE_HISTORY.key(),
                    MainShellTabId.OPERATOR_ACTION_LOG.key(),
                    MainShellTabId.IDENTITY_CHECK_HISTORY.key(),
                    MainShellTabId.ENV.key(),
                    MainShellTabId.GLOBAL_SETTINGS.key(),
                    MainShellTabId.USER_PROFILES.key(),
                    MainShellTabId.OPERATOR_USER_MANAGEMENT.key(),
                    MainShellTabId.PROCESSING_TREND.key());

    /**
     * All {@link MainShellTabId} keys except {@link MainShellTabId#TAB_ORGANIZER}: DEFAULT order then any enum-only
     * keys appended (new tabs at end).
     */
    public static List<String> completeFlatTabKeyOrder() {
        LinkedHashSet<String> keys = new LinkedHashSet<>(DEFAULT_FLAT_TAB_KEY_ORDER);
        for (MainShellTabId id : MainShellTabId.values()) {
            if (id != MainShellTabId.TAB_ORGANIZER) {
                keys.add(id.key());
            }
        }
        return List.copyOf(keys);
    }

    /** Default grouped layout when session has no {@code mainShellTabLayout}. */
    public static List<MainShellTabLayoutNode> groupedLayout() {
        List<MainShellTabLayoutNode> top = new ArrayList<>();
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.EQUIPMENT_STATUS_DASHBOARD.key(), "#994d66"));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.PROCESSING_TREND.key(), "#2f6fb3"));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.REMOTE_DESKTOP.key(), "#0000ff"));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.REQUEST_FORM_INPUT.key(), "#336633"));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.REQUEST_FORM_PIPELINE_CHECK.key(), "#9980e6"));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.COMPANY_CALENDAR.key(), "#800080"));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.MEMBER_ATTENDANCE.key(), "#800080"));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.MACHINE_CALENDAR.key(), "#800000"));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.MASTER_DISPATCH_SHEETS.key(), "#800000"));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.RUN.key(), "#0000ff"));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.PLAN_INPUT.key(), "#000080"));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.EQUIPMENT_GANTT_GRAPHIC.key(), "#1a4d1a"));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.DELIVERY_CALENDAR_VIEW.key(), "#669966"));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.OPERATOR_CARD.key(), "#336666"));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.CODE_LOOKUP_TABLES.key(), "#0000ff"));

        top.add(
                MainShellTabLayoutNode.groupNode(
                        "バッジ設定",
                        "#ffff00",
                        List.of(
                                MainShellTabLayoutNode.tabNode(MainShellTabId.UI_BADGE_DESIGN.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.PUSH_BUTTON_DESIGN.key(), ""),
                                MainShellTabLayoutNode.tabNode(
                                        MainShellTabId.GANTT_PERSON_BADGE_DESIGN.key(), ""),
                                MainShellTabLayoutNode.tabNode(
                                        MainShellTabId.REQUEST_FORM_PREVIEW_BADGE_DESIGN.key(), ""))));

        top.add(
                MainShellTabLayoutNode.groupNode(
                        "結果情報(デバッグ用)",
                        "#800080",
                        List.of(
                                MainShellTabLayoutNode.tabNode(MainShellTabId.PLAN_RESULT_VIEWER.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.STAGE1_PREVIEW.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.RESULT_DISPATCH.key(), ""))));

        top.add(
                MainShellTabLayoutNode.groupNode(
                        "その他",
                        "#666666",
                        List.of(
                                MainShellTabLayoutNode.tabNode(MainShellTabId.MASTER_SUMMARY.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.EXCLUDE_RULES.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.SPECIAL_RULES.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.PLAN_WORKSPACE_HISTORY.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.MEMORY_SETTINGS.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.API_MODEL_BENCHMARK.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.ACTUALS_STATUS.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.DAILY_REPORT_CSV_VIEW.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.PIPELINE_EXECUTION_TIMING.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.CACHE_HISTORY.key(), ""),
                                MainShellTabLayoutNode.tabNode(
                                        MainShellTabId.OPERATOR_ACTION_LOG.key(), ""),
                                MainShellTabLayoutNode.tabNode(
                                        MainShellTabId.IDENTITY_CHECK_HISTORY.key(), ""))));

        top.add(
                MainShellTabLayoutNode.groupNode(
                        "環境設定",
                        "#e64d4d",
                        List.of(
                                MainShellTabLayoutNode.tabNode(MainShellTabId.ENV.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.GLOBAL_SETTINGS.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.USER_PROFILES.key(), ""),
                                MainShellTabLayoutNode.tabNode(
                                        MainShellTabId.OPERATOR_USER_MANAGEMENT.key(), ""))));

        return List.copyOf(top);
    }
}
