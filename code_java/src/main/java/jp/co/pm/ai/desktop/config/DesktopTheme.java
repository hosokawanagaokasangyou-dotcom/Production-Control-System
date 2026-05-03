package jp.co.pm.ai.desktop.config;

import java.util.Locale;
import java.util.Objects;

import javafx.scene.Scene;

/**
 * UI color themes: default Modena (light) or palette overlays under {@code /css/theme-*.css}.
 */
public enum DesktopTheme {
    /** Default Modena + visible grid lines ({@code theme-light.css}). */
    LIGHT("light", "\u30e9\u30a4\u30c8", "theme-light.css"),
    DARK("dark", "\u30c0\u30fc\u30af", "theme-dark.css"),
    /** Cool light blue-gray workspace. */
    BLUE("blue", "\u30d6\u30eb\u30fc", "theme-blue.css"),
    /** Warm paper / reduced blue light. */
    SEPIA("sepia", "\u30bb\u30d4\u30a2", "theme-sepia.css"),
    /** Dark teal / blue-green. */
    OCEAN("ocean", "\u30aa\u30fc\u30b7\u30e3\u30f3", "theme-ocean.css"),
    /** Deep blue-black (high-contrast grid lines). */
    MIDNIGHT("midnight", "\u30df\u30c3\u30c9\u30ca\u30a4\u30c8", "theme-midnight.css"),
    /** Blue-gray dark (editor-style). */
    SLATE("slate", "\u30b9\u30ec\u30fc\u30c8", "theme-slate.css"),
    /** Warm dark with amber accent. */
    EMBER("ember", "\u30a8\u30f3\u30d0\u30fc", "theme-ember.css"),
    /** High-contrast light (readability). */
    CONTRAST("contrast", "\u30b3\u30f3\u30c8\u30e9\u30b9\u30c8", "theme-contrast.css");

    private static final String CSS_DIR = "/jp/co/pm/ai/desktop/css/";

    private final String id;
    private final String displayLabel;
    /** File name in {@link #CSS_DIR} (each theme includes grid-line contrast rules). */
    private final String overlayCssFile;

    DesktopTheme(String id, String displayLabel, String overlayCssFile) {
        this.id = id;
        this.displayLabel = displayLabel;
        this.overlayCssFile = overlayCssFile;
    }

    public String storedId() {
        return id;
    }

    public String displayLabel() {
        return displayLabel;
    }

    /** Dark palettes where log-row backgrounds should use higher-contrast tints. */
    public boolean isDarkUi() {
        return switch (this) {
            case DARK, MIDNIGHT, SLATE, EMBER, OCEAN -> true;
            case LIGHT, BLUE, SEPIA, CONTRAST -> false;
        };
    }

    public static DesktopTheme fromStored(String s) {
        if (s == null || s.isBlank()) {
            return LIGHT;
        }
        String key = s.trim().toLowerCase(Locale.ROOT);
        for (DesktopTheme t : values()) {
            if (t.id.equals(key)) {
                return t;
            }
        }
        return LIGHT;
    }

    public static DesktopTheme fromDisplayLabel(String label) {
        if (label == null) {
            return LIGHT;
        }
        for (DesktopTheme t : values()) {
            if (t.displayLabel.equals(label)) {
                return t;
            }
        }
        return LIGHT;
    }

    /**
     * Replaces any bundled {@code theme-*.css} overlay, then adds this theme's sheet at index 0 so
     * {@code pm-ai-desktop.css} can still override specifics.
     */
    public void applyTo(Scene scene) {
        if (scene == null) {
            return;
        }
        var sheets = scene.getStylesheets();
        sheets.removeIf(DesktopTheme::isBundledThemeOverlay);
        var url =
                Objects.requireNonNull(
                        DesktopTheme.class.getResource(CSS_DIR + overlayCssFile),
                        overlayCssFile);
        sheets.add(0, url.toExternalForm());
    }

    static boolean isBundledThemeOverlay(String url) {
        if (url == null) {
            return false;
        }
        return url.contains("/jp/co/pm/ai/desktop/css/theme-") && url.endsWith(".css");
    }
}
