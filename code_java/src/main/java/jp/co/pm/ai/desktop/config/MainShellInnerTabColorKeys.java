package jp.co.pm.ai.desktop.config;

import jp.co.pm.ai.desktop.MainShellTabId;

/**
 * 子タブ見出し色のセッションキー（{@link DesktopSessionState#innerTabHeaderColorByKey()}）。
 *
 * <p>区切りはタブ見出しに含まれない制御文字（{@code \u0001}）。
 */
public final class MainShellInnerTabColorKeys {

    private static final char SEP = '\u0001';

    private MainShellInnerTabColorKeys() {}

    public static String innerKey(MainShellTabId parent, String innerLabel) {
        if (parent == null) {
            return "";
        }
        return parent.key() + SEP + nz(innerLabel);
    }

    public static String nestedKey(
            MainShellTabId parent, String anchorInnerLabel, String nestedLabel) {
        if (parent == null) {
            return "";
        }
        return parent.key() + SEP + nz(anchorInnerLabel) + SEP + nz(nestedLabel);
    }

    private static String nz(String s) {
        return s != null ? s.strip() : "";
    }
}
