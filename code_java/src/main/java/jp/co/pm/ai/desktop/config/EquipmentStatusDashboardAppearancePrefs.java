package jp.co.pm.ai.desktop.config;

/**
 * ダッシュボード（設備現状）カードのレイアウト・タイポグラフィ・円グラフ見た目（セッション保存用）。
 */
public record EquipmentStatusDashboardAppearancePrefs(
        /** 0 = 幅に応じて自動折り返し、1〜12 = 固定列数。 */
        int columnCount,
        double cardWidth,
        /** 全画面時のカード幅倍率（%）。{@code cardWidth} に対する。 */
        double fullscreenCardWidthPercent,
        double cardPadding,
        double cardGapH,
        double cardGapV,
        double cardBorderRadius,
        /** {@code NONE}|{@code SUBTLE}|{@code MEDIUM}|{@code STRONG} */
        String cardShadowStyle,
        /** 空はシステム既定フォント。 */
        String fontFamily,
        double machineFontPx,
        double metaFontPx,
        double planFontPx,
        double pctFontPx,
        double chartSizePx,
        String chartDoneColorHex,
        String chartRemainColorHex,
        /** {@code FLAT}|{@code DEPTH}（立体風＝グラデーション＋影） */
        String chartStyle,
        boolean chartShadowEnabled,
        /** 全画面表示テーマ: {@code DARK}|{@code LIGHT}|{@code WALL} */
        String fullscreenTheme) {

    public static final String SHADOW_NONE = "NONE";
    public static final String SHADOW_SUBTLE = "SUBTLE";
    public static final String SHADOW_MEDIUM = "MEDIUM";
    public static final String SHADOW_STRONG = "STRONG";

    public static final String CHART_FLAT = "FLAT";
    public static final String CHART_DEPTH = "DEPTH";

    public static final String FULLSCREEN_THEME_DARK = "DARK";
    public static final String FULLSCREEN_THEME_LIGHT = "LIGHT";
    public static final String FULLSCREEN_THEME_WALL = "WALL";

    public EquipmentStatusDashboardAppearancePrefs {
        columnCount = clampInt(columnCount, 0, 12);
        cardWidth = clamp(cardWidth, 160, 520, 280);
        fullscreenCardWidthPercent = clamp(fullscreenCardWidthPercent, 80, 200, 121);
        cardPadding = clamp(cardPadding, 4, 32, 12);
        cardGapH = clamp(cardGapH, 0, 48, 12);
        cardGapV = clamp(cardGapV, 0, 48, 12);
        cardBorderRadius = clamp(cardBorderRadius, 0, 32, 8);
        cardShadowStyle = normalizeShadow(cardShadowStyle);
        fontFamily = fontFamily != null ? fontFamily.strip() : "";
        machineFontPx = clamp(machineFontPx, 9, 28, 15);
        metaFontPx = clamp(metaFontPx, 8, 20, 12);
        planFontPx = clamp(planFontPx, 8, 18, 11);
        pctFontPx = clamp(pctFontPx, 10, 32, 16);
        chartSizePx = clamp(chartSizePx, 40, 240, 96);
        chartDoneColorHex = normalizeHex(chartDoneColorHex, "#0d9488");
        chartRemainColorHex = normalizeHex(chartRemainColorHex, "#e2e8f0");
        chartStyle = normalizeChartStyle(chartStyle);
        fullscreenTheme = normalizeFullscreenTheme(fullscreenTheme);
    }

    public String fullscreenThemeStyleClass() {
        return "pm-equipment-status-fullscreen-theme-"
                + fullscreenTheme.toLowerCase(java.util.Locale.ROOT);
    }

    public static EquipmentStatusDashboardAppearancePrefs defaults() {
        return new EquipmentStatusDashboardAppearancePrefs(
                0,
                280,
                121,
                12,
                12,
                12,
                8,
                SHADOW_SUBTLE,
                "",
                15,
                12,
                11,
                16,
                96,
                "#0d9488",
                "#e2e8f0",
                CHART_FLAT,
                false,
                FULLSCREEN_THEME_DARK);
    }

    public double effectiveCardWidth(boolean fullscreen) {
        if (!fullscreen) {
            return cardWidth;
        }
        return cardWidth * fullscreenCardWidthPercent / 100.0;
    }

    private static double clamp(double v, double min, double max, double fallback) {
        if (!Double.isFinite(v)) {
            return fallback;
        }
        return Math.max(min, Math.min(max, v));
    }

    private static int clampInt(int v, int min, int max) {
        return Math.max(min, Math.min(max, v));
    }

    private static String normalizeShadow(String raw) {
        if (raw == null || raw.isBlank()) {
            return SHADOW_SUBTLE;
        }
        return switch (raw.strip().toUpperCase(java.util.Locale.ROOT)) {
            case SHADOW_NONE, SHADOW_SUBTLE, SHADOW_MEDIUM, SHADOW_STRONG -> raw.strip().toUpperCase(java.util.Locale.ROOT);
            default -> SHADOW_SUBTLE;
        };
    }

    private static String normalizeChartStyle(String raw) {
        if (raw == null || raw.isBlank()) {
            return CHART_FLAT;
        }
        return switch (raw.strip().toUpperCase(java.util.Locale.ROOT)) {
            case CHART_FLAT, CHART_DEPTH -> raw.strip().toUpperCase(java.util.Locale.ROOT);
            default -> CHART_FLAT;
        };
    }

    private static String normalizeFullscreenTheme(String raw) {
        if (raw == null || raw.isBlank()) {
            return FULLSCREEN_THEME_DARK;
        }
        return switch (raw.strip().toUpperCase(java.util.Locale.ROOT)) {
            case FULLSCREEN_THEME_DARK, FULLSCREEN_THEME_LIGHT, FULLSCREEN_THEME_WALL ->
                    raw.strip().toUpperCase(java.util.Locale.ROOT);
            default -> FULLSCREEN_THEME_DARK;
        };
    }

    private static String normalizeHex(String raw, String fallback) {
        if (raw == null || raw.isBlank()) {
            return fallback;
        }
        String s = raw.strip();
        if (!s.startsWith("#")) {
            s = "#" + s;
        }
        if (s.length() == 7 || s.length() == 9) {
            return s;
        }
        return fallback;
    }
}
