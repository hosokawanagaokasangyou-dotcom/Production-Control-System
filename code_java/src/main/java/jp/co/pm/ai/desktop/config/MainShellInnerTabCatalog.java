package jp.co.pm.ai.desktop.config;

import java.util.List;

import jp.co.pm.ai.desktop.MainShellTabId;

/**
 * メインシェル各タブ内の「子タブ」（{@link javafx.scene.control.TabPane}）見出しをタブ整理ツリーに載せるための一覧。
 *
 * <p>計画結果ビューアのように JSON 読込後にのみ現れるタブは代表ラベルで示す。
 */
public final class MainShellInnerTabCatalog {

    private MainShellInnerTabCatalog() {}

    /** @return 表示ラベル（永続 ID ではない） */
    public static List<String> labelsFor(MainShellTabId parent) {
        if (parent == null) {
            return List.of();
        }
        return switch (parent) {
            case DELIVERY_CALENDAR_VIEW ->
                    List.of(
                            "\u30ab\u30ec\u30f3\u30c0\u30fc",
                            "\u8a08\u753b\u6bd4\u8f03",
                            "\u30a2\u30e9\u30b8\u30f3\u52a0\u5de5\u8a08\u753b\u53d6\u5f97\u30c7\u30fc\u30bf");
            case DISPATCH_INTERACTIVE ->
                    List.of(
                            "\u30bf\u30b9\u30af\u00d7\u65e5\u4ed8",
                            "\u5de5\u7a0b+\u6a5f\u68b0\u00d7\u65e5");
            case PLAN_RESULT_VIEWER ->
                    List.of(
                            "\u751f\u7523\u8a08\u753b (multi_day) / \u30e1\u30f3\u30d0\u30fc\u52e4\u52d9",
                            "\uff08\u5404\u30c7\u30fc\u30bf\u30bb\u30c3\u30c8\uff09\u30b7\u30fc\u30c8",
                            "\u4e00\u89a7\uff08\u8868\uff09 / \u30ac\u30f3\u30c8");
            default -> List.of();
        };
    }
}
