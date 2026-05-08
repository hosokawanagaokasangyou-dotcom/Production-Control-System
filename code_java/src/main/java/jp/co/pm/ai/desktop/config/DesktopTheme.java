package jp.co.pm.ai.desktop.config;

import java.util.Locale;
import java.util.Objects;

import javafx.scene.Scene;

/**
 * UI color themes: default Modena (light) or palette overlays under {@code /css/theme-*.css}.
 */
public enum DesktopTheme {
    /** Default Modena + visible grid lines ({@code theme-light.css}). */
    LIGHT("light", "ライト", "theme-light.css"),
    DARK("dark", "ダーク", "theme-dark.css"),
    /** Cool light blue-gray workspace. */
    BLUE("blue", "ブルー", "theme-blue.css"),
    /** Warm paper / reduced blue light. */
    SEPIA("sepia", "セピア", "theme-sepia.css"),
    /** Dark teal / blue-green. */
    OCEAN("ocean", "オーシャン", "theme-ocean.css"),
    /** Deep blue-black (high-contrast grid lines). */
    MIDNIGHT("midnight", "ミッドナイト", "theme-midnight.css"),
    /** Blue-gray dark (editor-style). */
    SLATE("slate", "スレート", "theme-slate.css"),
    /** Warm dark with amber accent. */
    EMBER("ember", "エンバー", "theme-ember.css"),
    /** High-contrast light (readability). */
    CONTRAST("contrast", "コントラスト", "theme-contrast.css"),
    /** Deep blue-black with saturated blue accent. */
    MIDNIGHT_BLUE("midnight-blue", "ミッドナイトブルー", "theme-midnight-blue.css"),
    /** Deep navy / slate blue workspace. */
    DEEP_NAVY("deep-navy", "ディープネイビー", "theme-deep-navy.css"),
    /** Cool neutral dark gray (space gray). */
    SPACE_GRAY("space-gray", "スペースグレー", "theme-space-gray.css"),
    /** Warm charcoal with amber accent. */
    CHARCOAL("charcoal", "チャコール", "theme-charcoal.css"),
    /** Near-black ink with subtle blue accent. */
    INK_BLACK("ink-black", "インクブラック", "theme-ink-black.css"),
    /** Prussian / ink blue. */
    PRUSSIAN_BLUE("prussian-blue", "プルシアンブルー", "theme-prussian-blue.css"),
    /** Blue-gray gunmetal / steel blue. */
    GUNMETAL("gunmetal", "ガンメタル", "theme-gunmetal.css"),
    /** Nord-inspired frost accent on polar night. */
    NORD("nord", "ノード", "theme-nord.css"),
    /** Dracula-inspired purple / pink accent. */
    DRACULA("dracula", "ドラキュラ", "theme-dracula.css"),
    /** Deep purple aubergine. */
    AUBERGINE("aubergine", "オーベルジーヌ", "theme-aubergine.css"),
    /** Deep aquatic blue (oceanic). */
    OCEANIC("oceanic", "オセアニック", "theme-oceanic.css"),
    /** Cosmic deep blue-violet. */
    DEEP_ASTRAL("deep-astral", "ディープアストラル", "theme-deep-astral.css"),
    /** Material Design dark (purple accent). */
    MATERIAL_DARK("material-dark", "マテリアルダーク", "theme-material-dark.css"),
    /** True black background tuned for OLED. */
    OLED_BLACK("oled-black", "オーレッドブラック", "theme-oled-black.css"),
    /** Dark forest / moss green tint. */
    NIGHT_MOSS("night-moss", "ナイトモス", "theme-night-moss.css"),
    /** Industrial steel gray. */
    STEEL_GRAY("steel-gray", "スチールグレー", "theme-steel-gray.css"),
    /** Oxford / academic navy. */
    OXFORD_BLUE("oxford-blue", "オックスフォードブルー", "theme-oxford-blue.css");

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
            case LIGHT, BLUE, SEPIA, CONTRAST -> false;
            default -> true;
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
