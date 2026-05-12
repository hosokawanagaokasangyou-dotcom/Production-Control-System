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
 */
public final class MainShellTabLayoutDefaults {

    private MainShellTabLayoutDefaults() {}

    /** Flat tab key order (reset-flat button and merge order for missing keys). */
    public static final List<String> DEFAULT_FLAT_TAB_KEY_ORDER =
            List.of(
                    MainShellTabId.RUN.key(),
                    MainShellTabId.PLAN_INPUT.key(),
                    MainShellTabId.DISPATCH_INTERACTIVE.key(),
                    MainShellTabId.EQUIPMENT_GANTT_GRAPHIC.key(),
                    MainShellTabId.DELIVERY_CALENDAR_VIEW.key(),
                    MainShellTabId.OPERATOR_CARD.key(),
                    MainShellTabId.UI_BADGE_DESIGN.key(),
                    MainShellTabId.PUSH_BUTTON_DESIGN.key(),
                    MainShellTabId.GANTT_PERSON_BADGE_DESIGN.key(),
                    MainShellTabId.ENV.key(),
                    MainShellTabId.MEMORY_SETTINGS.key(),
                    MainShellTabId.GLOBAL_SETTINGS.key(),
                    MainShellTabId.PLAN_RESULT_VIEWER.key(),
                    MainShellTabId.STAGE1_PREVIEW.key(),
                    MainShellTabId.RESULT_DISPATCH.key(),
                    MainShellTabId.MASTER_SUMMARY.key(),
                    MainShellTabId.EXCLUDE_RULES.key(),
                    MainShellTabId.SPECIAL_RULES.key(),
                    MainShellTabId.ACTUALS_STATUS.key(),
                    MainShellTabId.USER_PROFILES.key(),
                    MainShellTabId.PLAN_WORKSPACE_HISTORY.key());

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
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.RUN.key(), ""));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.PLAN_INPUT.key(), ""));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.DISPATCH_INTERACTIVE.key(), ""));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.EQUIPMENT_GANTT_GRAPHIC.key(), ""));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.DELIVERY_CALENDAR_VIEW.key(), ""));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.OPERATOR_CARD.key(), ""));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.PLAN_WORKSPACE_HISTORY.key(), ""));

        top.add(
                MainShellTabLayoutNode.groupNode(
                        "\u30d0\u30c3\u30b8\u8a2d\u5b9a",
                        "",
                        List.of(
                                MainShellTabLayoutNode.tabNode(MainShellTabId.UI_BADGE_DESIGN.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.PUSH_BUTTON_DESIGN.key(), ""),
                                MainShellTabLayoutNode.tabNode(
                                        MainShellTabId.GANTT_PERSON_BADGE_DESIGN.key(), ""))));

        top.add(
                MainShellTabLayoutNode.groupNode(
                        "\u74b0\u5883\u8a2d\u5b9a",
                        "",
                        List.of(
                                MainShellTabLayoutNode.tabNode(MainShellTabId.ENV.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.MEMORY_SETTINGS.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.GLOBAL_SETTINGS.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.USER_PROFILES.key(), ""))));

        top.add(
                MainShellTabLayoutNode.groupNode(
                        "\u7d50\u679c\u60c5\u5831",
                        "",
                        List.of(
                                MainShellTabLayoutNode.tabNode(MainShellTabId.PLAN_RESULT_VIEWER.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.STAGE1_PREVIEW.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.RESULT_DISPATCH.key(), ""))));

        top.add(
                MainShellTabLayoutNode.groupNode(
                        "\u305d\u306e\u4ed6",
                        "",
                        List.of(
                                MainShellTabLayoutNode.tabNode(MainShellTabId.MASTER_SUMMARY.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.EXCLUDE_RULES.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.SPECIAL_RULES.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.ACTUALS_STATUS.key(), ""))));

        return List.copyOf(top);
    }
}
