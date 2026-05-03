package jp.co.pm.ai.desktop.config;

import java.util.Objects;

import javafx.scene.Scene;

/**
 * UI color theme (Modena light vs dark palette overlay).
 */
public enum DesktopTheme {
    LIGHT("light", "\u30e9\u30a4\u30c8"),
    DARK("dark", "\u30c0\u30fc\u30af");

    private final String id;
    private final String displayLabel;

    DesktopTheme(String id, String displayLabel) {
        this.id = id;
        this.displayLabel = displayLabel;
    }

    public String storedId() {
        return id;
    }

    public String displayLabel() {
        return displayLabel;
    }

    public static DesktopTheme fromStored(String s) {
        if (s == null || s.isBlank()) {
            return LIGHT;
        }
        if ("dark".equalsIgnoreCase(s.trim())) {
            return DARK;
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
     * Applies this theme to the scene stylesheet list. Keeps {@code pm-ai-desktop.css} and other app sheets;
     * adds or removes only {@code theme-dark.css}.
     */
    public void applyTo(Scene scene) {
        if (scene == null) {
            return;
        }
        var sheets = scene.getStylesheets();
        sheets.removeIf(DesktopTheme::isDarkThemeUrl);
        if (this == DARK) {
            var url = Objects.requireNonNull(
                    DesktopTheme.class.getResource("/jp/co/pm/ai/desktop/css/theme-dark.css"),
                    "theme-dark.css");
            sheets.add(0, url.toExternalForm());
        }
    }

    private static boolean isDarkThemeUrl(String url) {
        return url != null && (url.endsWith("theme-dark.css") || url.contains("/theme-dark.css"));
    }
}
