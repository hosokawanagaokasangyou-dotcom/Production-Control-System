package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.file.Path;
import java.util.Locale;
import java.util.Map;

import jp.co.pm.ai.desktop.config.AppPaths;

/** mstsc 起動前に .rdp へ解像度・全画面／ウィンドウ設定を反映する。 */
public final class RdpLaunchDisplaySettings {

    public static final int DEFAULT_WIDTH = 1280;
    public static final int DEFAULT_HEIGHT = 800;

    private RdpLaunchDisplaySettings() {}

    /** {@code true} = 全画面（screen mode id 2）。 */
    public static boolean resolveFullScreen(Map<String, String> ui) {
        String raw = trimFromUi(ui, AppPaths.KEY_PM_AI_RDP_FULLSCREEN);
        if (raw.isEmpty()) {
            return false;
        }
        String v = raw.toLowerCase(Locale.ROOT);
        return "1".equals(v) || "true".equals(v) || "on".equals(v) || "yes".equals(v);
    }

    public static int resolveWidth(Map<String, String> ui) {
        return parsePositiveInt(trimFromUi(ui, AppPaths.KEY_PM_AI_RDP_DESKTOP_WIDTH), DEFAULT_WIDTH);
    }

    public static int resolveHeight(Map<String, String> ui) {
        return parsePositiveInt(trimFromUi(ui, AppPaths.KEY_PM_AI_RDP_DESKTOP_HEIGHT), DEFAULT_HEIGHT);
    }

    /**
     * 起動直前に .rdp へ表示設定を書き込む。
     *
     * @return 署名行を削除した場合 {@code true}
     */
    public static boolean applyToProfile(Path rdpProfile, Map<String, String> ui) throws IOException {
        int width = resolveWidth(ui);
        int height = resolveHeight(ui);
        boolean fullScreen = resolveFullScreen(ui);
        return RdpProfileEditor.applyDesktopDisplay(rdpProfile, width, height, fullScreen);
    }

    public static String formatSummary(Map<String, String> ui) {
        if (resolveFullScreen(ui)) {
            return "全画面";
        }
        return resolveWidth(ui) + " x " + resolveHeight(ui) + "（ウィンドウ）";
    }

    private static int parsePositiveInt(String raw, int defaultValue) {
        if (raw == null || raw.isBlank()) {
            return defaultValue;
        }
        try {
            int value = Integer.parseInt(raw.trim());
            return value > 0 ? value : defaultValue;
        } catch (NumberFormatException ex) {
            return defaultValue;
        }
    }

    private static String trimFromUi(Map<String, String> ui, String key) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String raw = u.get(key);
        if (raw == null || raw.isBlank()) {
            raw = System.getenv(key);
        }
        return raw != null ? raw.trim() : "";
    }
}
