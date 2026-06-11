package jp.co.pm.ai.desktop.io;

import java.util.Locale;
import java.util.Map;

import jp.co.pm.ai.desktop.config.AppPaths;

/** リモートデスクトップタブ内 mstsc 埋め込み設定。 */
public final class RdpEmbedSettings {

    /** 1/true/on で右ペイン上部へ埋め込み（既定 ON）。0/false/off で従来の別ウィンドウ。 */
    public static final String KEY_PM_AI_RDP_EMBED_IN_TAB = "PM_AI_RDP_EMBED_IN_TAB";

    private RdpEmbedSettings() {}

    public static boolean isEmbedInTabEnabled(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String raw = u.get(KEY_PM_AI_RDP_EMBED_IN_TAB);
        if (raw == null || raw.isBlank()) {
            raw = System.getenv(KEY_PM_AI_RDP_EMBED_IN_TAB);
        }
        if (raw == null || raw.isBlank()) {
            return true;
        }
        String v = raw.trim().toLowerCase(Locale.ROOT);
        return !("0".equals(v) || "false".equals(v) || "off".equals(v) || "no".equals(v));
    }
}
