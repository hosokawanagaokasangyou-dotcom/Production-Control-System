package jp.co.pm.ai.desktop.config;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

class EquipmentStatusDashboardAppearancePrefsTest {

    @Test
    void defaults_areStable() {
        EquipmentStatusDashboardAppearancePrefs d = EquipmentStatusDashboardAppearancePrefs.defaults();
        Assertions.assertEquals(0, d.columnCount());
        Assertions.assertEquals(280, d.cardWidth(), 0.01);
        Assertions.assertEquals(96, d.chartSizePx(), 0.01);
        Assertions.assertEquals(
                EquipmentStatusDashboardAppearancePrefs.FULLSCREEN_THEME_DARK, d.fullscreenTheme());
    }

    @Test
    void clampsOutOfRangeValues() {
        EquipmentStatusDashboardAppearancePrefs p =
                new EquipmentStatusDashboardAppearancePrefs(
                        99,
                        50,
                        50,
                        1,
                        -5,
                        100,
                        -1,
                        "INVALID",
                        "  ",
                        5,
                        50,
                        50,
                        50,
                        10,
                        "bad",
                        "",
                        "X",
                        true,
                        "BAD");
        Assertions.assertEquals(12, p.columnCount());
        Assertions.assertEquals(160, p.cardWidth(), 0.01);
        Assertions.assertEquals(80, p.fullscreenCardWidthPercent(), 0.01);
        Assertions.assertEquals(EquipmentStatusDashboardAppearancePrefs.SHADOW_SUBTLE, p.cardShadowStyle());
        Assertions.assertEquals(EquipmentStatusDashboardAppearancePrefs.CHART_FLAT, p.chartStyle());
        Assertions.assertEquals("#0d9488", p.chartDoneColorHex());
        Assertions.assertEquals(
                EquipmentStatusDashboardAppearancePrefs.FULLSCREEN_THEME_DARK, p.fullscreenTheme());
    }

    @Test
    void effectiveCardWidth_scalesForFullscreen() {
        EquipmentStatusDashboardAppearancePrefs p =
                new EquipmentStatusDashboardAppearancePrefs(
                        0,
                        200,
                        150,
                        12,
                        12,
                        12,
                        8,
                        EquipmentStatusDashboardAppearancePrefs.SHADOW_SUBTLE,
                        "",
                        15,
                        12,
                        11,
                        16,
                        96,
                        "#0d9488",
                        "#e2e8f0",
                        EquipmentStatusDashboardAppearancePrefs.CHART_FLAT,
                        false,
                        EquipmentStatusDashboardAppearancePrefs.FULLSCREEN_THEME_LIGHT);
        Assertions.assertEquals(200, p.effectiveCardWidth(false), 0.01);
        Assertions.assertEquals(300, p.effectiveCardWidth(true), 0.01);
        Assertions.assertEquals(
                EquipmentStatusDashboardAppearancePrefs.FULLSCREEN_THEME_LIGHT, p.fullscreenTheme());
    }

    @Test
    void normalizeHex_rejectsNonHexAndNormalizesCase() {
        Assertions.assertEquals("#0d9488", prefsWithColors("#zzzzzz", "#e2e8f0").chartDoneColorHex());
        Assertions.assertEquals("#0d9488", prefsWithColors("0D9488", "#e2e8f0").chartDoneColorHex());
        Assertions.assertEquals("#ff000080", prefsWithColors("#FF000080", "#e2e8f0").chartDoneColorHex());
        Assertions.assertEquals("#e2e8f0", prefsWithColors("#0d9488", "#12345").chartRemainColorHex());
    }

    @Test
    void fontSizes_haveReadableLowerBounds() {
        EquipmentStatusDashboardAppearancePrefs tiny =
                new EquipmentStatusDashboardAppearancePrefs(
                        0,
                        280,
                        121,
                        12,
                        12,
                        12,
                        8,
                        EquipmentStatusDashboardAppearancePrefs.SHADOW_SUBTLE,
                        "",
                        1,
                        1,
                        1,
                        1,
                        96,
                        "#0d9488",
                        "#e2e8f0",
                        EquipmentStatusDashboardAppearancePrefs.CHART_FLAT,
                        false,
                        EquipmentStatusDashboardAppearancePrefs.FULLSCREEN_THEME_DARK);
        Assertions.assertEquals(10, tiny.machineFontPx(), 0.01);
        Assertions.assertEquals(10, tiny.metaFontPx(), 0.01);
        Assertions.assertEquals(10, tiny.planFontPx(), 0.01);
        Assertions.assertEquals(10, tiny.pctFontPx(), 0.01);
    }

    private static EquipmentStatusDashboardAppearancePrefs prefsWithColors(String done, String remain) {
        return new EquipmentStatusDashboardAppearancePrefs(
                0,
                280,
                121,
                12,
                12,
                12,
                8,
                EquipmentStatusDashboardAppearancePrefs.SHADOW_SUBTLE,
                "",
                15,
                12,
                11,
                16,
                96,
                done,
                remain,
                EquipmentStatusDashboardAppearancePrefs.CHART_FLAT,
                false,
                EquipmentStatusDashboardAppearancePrefs.FULLSCREEN_THEME_DARK);
    }

    @Test
    void fullscreenThemeStyleClass_mapsKeyToCss() {
        EquipmentStatusDashboardAppearancePrefs wall =
                new EquipmentStatusDashboardAppearancePrefs(
                        0,
                        280,
                        121,
                        12,
                        12,
                        12,
                        8,
                        EquipmentStatusDashboardAppearancePrefs.SHADOW_SUBTLE,
                        "",
                        15,
                        12,
                        11,
                        16,
                        96,
                        "#0d9488",
                        "#e2e8f0",
                        EquipmentStatusDashboardAppearancePrefs.CHART_FLAT,
                        false,
                        EquipmentStatusDashboardAppearancePrefs.FULLSCREEN_THEME_WALL);
        Assertions.assertEquals(
                "pm-equipment-status-fullscreen-theme-wall", wall.fullscreenThemeStyleClass());
    }
}
