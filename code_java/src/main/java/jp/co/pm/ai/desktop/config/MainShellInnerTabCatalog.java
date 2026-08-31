package jp.co.pm.ai.desktop.config;

import java.util.List;

import jp.co.pm.ai.desktop.MainShellTabId;

/**
 * Catalog of TabPane child labels and optional {@link javafx.scene.control.TitledPane} rows under a child tab,
 * for the tab organizer tree (see {@code MainShellTabOrganizerTabController}).
 *
 * <p>Rulebook: {@code .cursor/rules/main-shell-tab-management.mdc}
 */
public final class MainShellInnerTabCatalog {

    private MainShellInnerTabCatalog() {}

    /**
     * TitledPane headings under the inner tab at {@code innerTabIndex} in {@link #labelsFor} order (0-based).
     */
    public static List<String> titledPaneLabelsUnderInnerTab(
            MainShellTabId parent, int innerTabIndex) {
        if (parent == MainShellTabId.DELIVERY_CALENDAR_VIEW && innerTabIndex == 0) {
            return List.of(
                    "\u64cd\u4f5c\u30fb\u30bd\u30fc\u30b9",
                    "\u30c7\u30fc\u30bf\u8868");
        }
        if (parent == MainShellTabId.DELIVERY_CALENDAR_VIEW
                && (innerTabIndex == 1
                        || innerTabIndex == 2
                        || innerTabIndex == 3
                        || innerTabIndex == 4)) {
            return List.of(
                    "\u64cd\u4f5c\u30fb\u30bd\u30fc\u30b9",
                    "\u30c7\u30fc\u30bf\u8868");
        }
        return List.of();
    }

    /**
     * 子タブ直下のさらに内側 TabPane の見出し（{@code innerTabIndex} は {@link #labelsFor} 順・0 始まり）。
     */
    public static List<String> nestedInnerTabLabelsUnderInnerTab(
            MainShellTabId parent, int innerTabIndex) {
        // 依頼書入力「マスター一覧」内 TabPane（ReconciliationApp）
        if (parent == MainShellTabId.REQUEST_FORM_INPUT && innerTabIndex == 4) {
            return List.of("機械コード", "工程マスタ", "加工内容マスタ");
        }
        return List.of();
    }

    /** Display labels for TabPane tabs (not persisted IDs). */
    public static List<String> labelsFor(MainShellTabId parent) {
        if (parent == null) {
            return List.of();
        }
        return switch (parent) {
            case CODE_LOOKUP_TABLES ->
                    List.of(
                            "使用原反→ロール長(m)",
                            "製品名→ロール長(m)",
                            "製品名→製品幅(mm)",
                            "製品名→厚み(mm)",
                            "製品名→製品長(mm)",
                            "使用原反→原反幅(mm)",
                            "リポジトリから上書き");
                    case MASTER_DISPATCH_SHEETS -> List.of("資格（skills）", "必要人数（need）", "加工速度（speed）", "組み合わせ表");
            case DELIVERY_CALENDAR_VIEW ->
                    List.of(
                            "\u30a2\u30e9\u30fb\u5b9f\u7e3e\u30fb\u30b7\u30b9\u6bd4\u8f03",
                            "\u914d\u53f0\u7d50\u679c",
                            "\u914d\u53f0\u7d50\u679c\uff08\u30bf\u30b9\u30af\u96c6\u7d04\uff09",
                            "\u52a0\u5de5\u5b9f\u7e3e",
                            "\u30a2\u30e9\u30b8\u30f3\u52a0\u5de5\u8a08\u753b\u53d6\u5f97\u30c7\u30fc\u30bf");
            case ENV ->
                    List.of(
                            "\u74b0\u5883\u5909\u6570\u4e00\u89a7",
                            "\u914d\u53f0 Gemini \u30e2\u30c7\u30eb\u512a\u5148");
            case PLAN_RESULT_VIEWER ->
                    List.of(
                            "\u751f\u7523\u8a08\u753b (multi_day) / \u30e1\u30f3\u30d0\u30fc\u52e4\u52d9",
                            "\uff08\u5404\u30c7\u30fc\u30bf\u30bb\u30c3\u30c8\uff09\u30b7\u30fc\u30c8",
                            "\u4e00\u89a7\uff08\u8868\uff09 / \u30ac\u30f3\u30c8");
            case SPECIAL_RULES ->
                    List.of(
                            "要約",
                            "列挙",
                            "ルールビルダー",
                            "ルール試走",
                            "適用トレース",
                            "工程優先",
                            "JSON");
            case REQUEST_FORM_INPUT ->
                    List.of(
                            "一括照合データベース・受注管理",
                            "目次シート",
                            "【設定】",
                            "後加工商品マスタ",
                            "マスター一覧");
            default -> List.of();
        };
    }
}
