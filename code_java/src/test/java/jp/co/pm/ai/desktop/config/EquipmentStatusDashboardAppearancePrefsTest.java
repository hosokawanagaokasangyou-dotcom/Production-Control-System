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
