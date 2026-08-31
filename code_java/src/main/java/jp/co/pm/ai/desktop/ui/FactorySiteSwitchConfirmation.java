package jp.co.pm.ai.desktop.ui;

import jp.co.pm.ai.desktop.config.FactorySite;

/**
 * 工場切替（湖南／国分）のユーザー確認。起動時の工場復元では出さない。
 */
public final class FactorySiteSwitchConfirmation {

    public static final String TITLE = "工場切替の確認";

    private FactorySiteSwitchConfirmation() {}

    /** ツールバー／グローバル設定からの切替は true。起動時復元は false。 */
    public static boolean shouldPromptUser(boolean startup) {
        return !startup;
    }

    public static String contentText(FactorySite from, FactorySite to) {
        String fromLabel = from != null ? from.displayLabelJa() : "（未設定）";
        String toLabel = to != null ? to.displayLabelJa() : "";
        return fromLabel + "から" + toLabel + "へ切り替えます。よろしいですか？";
    }
}
