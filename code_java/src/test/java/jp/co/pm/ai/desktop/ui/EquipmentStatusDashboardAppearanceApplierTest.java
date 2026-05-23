package jp.co.pm.ai.desktop.ui;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.config.EquipmentStatusDashboardAppearancePrefs;

class EquipmentStatusDashboardAppearanceApplierTest {

    @Test
    void fixedColumnWrapInnerWidth_computesFromCardWidthAndGaps() {
        EquipmentStatusDashboardAppearancePrefs prefs =
                new EquipmentStatusDashboardAppearancePrefs(
                        3,
                        200,
                        100,
                        12,
                        10,
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
                        EquipmentStatusDashboardAppearancePrefs.FULLSCREEN_THEME_DARK);
        Assertions.assertEquals(620, EquipmentStatusDashboardAppearanceApplier.fixedColumnWrapInnerWidth(prefs, false), 0.01);
        Assertions.assertEquals(620, EquipmentStatusDashboardAppearanceApplier.fixedColumnWrapInnerWidth(prefs, true), 0.01);
        Assertions.assertFalse(EquipmentStatusDashboardAppearanceApplier.scrollShouldFitToWidth(prefs));
    }

    @Test
    void scrollShouldFitToWidth_isTrueForAutoColumns() {
        Assertions.assertTrue(
                EquipmentStatusDashboardAppearanceApplier.scrollShouldFitToWidth(
                        EquipmentStatusDashboardAppearancePrefs.defaults()));
    }

    @Test
    void fixedColumnWrapInnerWidth_returnsNegativeForAutoLayout() {
        EquipmentStatusDashboardAppearancePrefs auto = EquipmentStatusDashboardAppearancePrefs.defaults();
        Assertions.assertTrue(EquipmentStatusDashboardAppearanceApplier.usesAutoColumnLayout(auto));
        Assertions.assertEquals(
                -1, EquipmentStatusDashboardAppearanceApplier.fixedColumnWrapInnerWidth(auto, false), 0.01);
    }
}
