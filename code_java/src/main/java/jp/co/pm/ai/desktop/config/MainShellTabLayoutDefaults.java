package jp.co.pm.ai.desktop.config;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;

import jp.co.pm.ai.desktop.MainShellTabId;

/**
 * メインシェルタブの既定の入れ子構成（タブ整理の初期状態・セッション未保存時）。
 *
 * <p>新規 {@link MainShellTabId} が追加された場合は {@link #DEFAULT_FLAT_TAB_KEY_ORDER} の末尾にキーを足す（トップレベル最後＝タブ整理の直前）。
 */
public final class MainShellTabLayoutDefaults {

    private MainShellTabLayoutDefaults() {}

    /**
     * フラット一列に並べるときのキー順（「フラット初期構成に戻す」およびマージ時の欠落キー挿入順）。
     */
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
                    MainShellTabId.ACTUALS_STATUS.key());

    /**
     * フラット復帰および「欠けたキー」のマージ順に、{@link MainShellTabId} のうち {@link
     * MainShellTabId#TAB_ORGANIZER} 以外をすべて含める（コード追加タブは末尾）。
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

    /**
     * 既定のグループ付き構成（セッションに mainShellTabLayout が無いときの初期適用）。
     */
    public static List<MainShellTabLayoutNode> groupedLayout() {
        List<MainShellTabLayoutNode> top = new ArrayList<>();
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.RUN.key(), ""));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.PLAN_INPUT.key(), ""));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.DISPATCH_INTERACTIVE.key(), ""));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.EQUIPMENT_GANTT_GRAPHIC.key(), ""));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.DELIVERY_CALENDAR_VIEW.key(), ""));
        top.add(MainShellTabLayoutNode.tabNode(MainShellTabId.OPERATOR_CARD.key(), ""));

        top.add(
                MainShellTabLayoutNode.groupNode(
                        "バッジ設定",
                        "",
                        List.of(
                                MainShellTabLayoutNode.tabNode(MainShellTabId.UI_BADGE_DESIGN.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.PUSH_BUTTON_DESIGN.key(), ""),
                                MainShellTabLayoutNode.tabNode(
                                        MainShellTabId.GANTT_PERSON_BADGE_DESIGN.key(), ""))));

        top.add(
                MainShellTabLayoutNode.groupNode(
                        "環境設定",
                        "",
                        List.of(
                                MainShellTabLayoutNode.tabNode(MainShellTabId.ENV.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.MEMORY_SETTINGS.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.GLOBAL_SETTINGS.key(), ""))));

        top.add(
                MainShellTabLayoutNode.groupNode(
                        "結果情報",
                        "",
                        List.of(
                                MainShellTabLayoutNode.tabNode(MainShellTabId.PLAN_RESULT_VIEWER.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.STAGE1_PREVIEW.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.RESULT_DISPATCH.key(), ""))));

        top.add(
                MainShellTabLayoutNode.groupNode(
                        "その他",
                        "",
                        List.of(
                                MainShellTabLayoutNode.tabNode(MainShellTabId.MASTER_SUMMARY.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.EXCLUDE_RULES.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.SPECIAL_RULES.key(), ""),
                                MainShellTabLayoutNode.tabNode(MainShellTabId.ACTUALS_STATUS.key(), ""))));

        return List.copyOf(top);
    }
}
