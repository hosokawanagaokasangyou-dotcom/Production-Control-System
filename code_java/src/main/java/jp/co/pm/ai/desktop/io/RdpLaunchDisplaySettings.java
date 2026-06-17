package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

import jp.co.pm.ai.desktop.config.AppPaths;

/** mstsc 起動前に .rdp へ解像度・全画面／ウィンドウ設定を反映する。 */
public final class RdpLaunchDisplaySettings {

    public static final int DEFAULT_WIDTH = 1920;
    public static final int DEFAULT_HEIGHT = 1080;
    public static final int MIN_WIDTH = 270;
    public static final int MIN_HEIGHT = 200;
    public static final int MAX_WIDTH = 3840;
    public static final int MAX_HEIGHT = 2160;

    /** 起動時に確定した表示設定。 */
    public record LaunchDisplay(boolean fullScreen, int width, int height) {

        public String summaryText() {
            if (fullScreen) {
                return "全画面";
            }
            return width + " x " + height + "（ウィンドウ）";
        }
    }

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
        return clampWidth(parsePositiveInt(trimFromUi(ui, AppPaths.KEY_PM_AI_RDP_DESKTOP_WIDTH), DEFAULT_WIDTH));
    }

    public static int resolveHeight(Map<String, String> ui) {
        return clampHeight(parsePositiveInt(trimFromUi(ui, AppPaths.KEY_PM_AI_RDP_DESKTOP_HEIGHT), DEFAULT_HEIGHT));
    }

    public static int clampWidth(int width) {
        return Math.max(MIN_WIDTH, Math.min(MAX_WIDTH, width));
    }

    public static int clampHeight(int height) {
        return Math.max(MIN_HEIGHT, Math.min(MAX_HEIGHT, height));
    }

    /**
     * 起動プロファイルと環境変数から表示設定を確定する。
     *
     * @param profile 選択中プロファイル（行 UI 由来を推奨）
     * @param ui 環境変数タブ由来のマップ（プロファイル未設定フィールドのフォールバック）
     */
    public static LaunchDisplay resolveLaunchDisplay(
            RdpLaunchProfile profile, Map<String, String> ui) {
        Map<String, String> env = ui != null ? ui : Map.of();
        RdpLaunchProfile p = profile != null ? profile : RdpLaunchProfile.empty(1);

        boolean fullScreen;
        if (p.fullScreen() != null) {
            fullScreen = p.fullScreen();
        } else {
            fullScreen = resolveFullScreen(env);
        }

        int width;
        if (p.desktopWidth() != null) {
            width = clampWidth(p.desktopWidth());
        } else {
            width = resolveWidth(env);
        }

        int height;
        if (p.desktopHeight() != null) {
            height = clampHeight(p.desktopHeight());
        } else {
            height = resolveHeight(env);
        }

        return new LaunchDisplay(fullScreen, width, height);
    }

    /** {@link LaunchDisplay} を ui マップへ書き込む（起動直前用）。 */
    public static Map<String, String> applyLaunchDisplayToUi(
            Map<String, String> ui, LaunchDisplay display) {
        Map<String, String> merged = new HashMap<>(ui != null ? ui : Map.of());
        merged.put(AppPaths.KEY_PM_AI_RDP_FULLSCREEN, display.fullScreen() ? "1" : "0");
        merged.put(AppPaths.KEY_PM_AI_RDP_DESKTOP_WIDTH, String.valueOf(display.width()));
        merged.put(AppPaths.KEY_PM_AI_RDP_DESKTOP_HEIGHT, String.valueOf(display.height()));
        return merged;
    }

    /**
     * 起動直前に .rdp へ表示設定を書き込む。
     *
     * @return 署名行を削除した場合 {@code true}
     */
    public static boolean applyToProfile(Path rdpProfile, Map<String, String> ui) throws IOException {
        LaunchDisplay display = resolveLaunchDisplay(null, ui);
        return applyToProfile(rdpProfile, display);
    }

    public static boolean applyToProfile(Path rdpProfile, LaunchDisplay display) throws IOException {
        return RdpProfileEditor.applyDesktopDisplay(
                rdpProfile, display.width(), display.height(), display.fullScreen());
    }

    public static String formatSummary(Map<String, String> ui) {
        return resolveLaunchDisplay(null, ui).summaryText();
    }

    public static String formatSummary(LaunchDisplay display) {
        return display.summaryText();
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
