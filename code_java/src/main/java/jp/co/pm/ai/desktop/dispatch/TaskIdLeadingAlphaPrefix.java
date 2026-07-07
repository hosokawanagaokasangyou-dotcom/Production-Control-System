package jp.co.pm.ai.desktop.dispatch;

import java.util.Locale;

/**
 * 依頼NO（タスクID）先頭の英字 1〜2 文字を接頭辞タブ用キーとして抽出する。
 *
 * <p>例: {@code JR260703} → {@code JR}, {@code C7-10} → {@code C}, {@code GB6064} → {@code GB}
 */
public final class TaskIdLeadingAlphaPrefix {

    public static final String OTHER = "（その他）";

    private TaskIdLeadingAlphaPrefix() {}

    public static String extract(String taskId) {
        if (taskId == null || taskId.isBlank()) {
            return OTHER;
        }
        String t = taskId.strip().toUpperCase(Locale.ROOT);
        if (t.length() >= 2
                && Character.isLetter(t.charAt(0))
                && Character.isLetter(t.charAt(1))) {
            return t.substring(0, 2);
        }
        if (!t.isEmpty() && Character.isLetter(t.charAt(0))) {
            return t.substring(0, 1);
        }
        return OTHER;
    }
}
