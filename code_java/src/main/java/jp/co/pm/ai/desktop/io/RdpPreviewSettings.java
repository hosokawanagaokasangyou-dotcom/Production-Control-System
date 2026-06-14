package jp.co.pm.ai.desktop.io;

import java.util.Locale;
import java.util.Map;

/** リモートデスクトップタブ内の読み取り専用プレビュー設定。 */
public final class RdpPreviewSettings {

    /** 1/true/on で右ペイン上部に低 fps プレビュー（既定 ON）。失敗時は別ウィンドウのまま。 */
    public static final String KEY_PM_AI_RDP_PREVIEW_IN_TAB = "PM_AI_RDP_PREVIEW_IN_TAB";

    private RdpPreviewSettings() {}

    public static boolean isPreviewInTabEnabled(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String raw = u.get(KEY_PM_AI_RDP_PREVIEW_IN_TAB);
        if (raw == null || raw.isBlank()) {
            raw = System.getenv(KEY_PM_AI_RDP_PREVIEW_IN_TAB);
        }
        if (raw == null || raw.isBlank()) {
            return true;
        }
        String v = raw.trim().toLowerCase(Locale.ROOT);
        return !("0".equals(v) || "false".equals(v) || "off".equals(v) || "no".equals(v));
    }
}
