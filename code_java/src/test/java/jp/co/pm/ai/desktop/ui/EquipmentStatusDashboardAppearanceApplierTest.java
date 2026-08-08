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

    @Test
    void computeFlowSpec_autoColumnsUseFullViewportInnerWidth() {
        EquipmentStatusDashboardAppearancePrefs auto = EquipmentStatusDashboardAppearancePrefs.defaults();
        EquipmentStatusDashboardAppearanceApplier.FlowLayoutSpec spec =
                EquipmentStatusDashboardAppearanceApplier.computeFlowSpec(auto, false, 1200, 12, 12);
        Assertions.assertTrue(spec.fillViewport());
        Assertions.assertEquals(1176, spec.wrapLength(), 0.01);
        Assertions.assertEquals(-1, spec.totalWidth(), 0.01);
        int columns =
                (int) Math.floor((spec.wrapLength() + auto.cardGapH()) / (auto.cardWidth() + auto.cardGapH()));
        Assertions.assertEquals(4, columns);
    }

    @Test
    void computeFlowSpec_autoColumnsFallBackToCardWidthBeforeFirstLayout() {
        EquipmentStatusDashboardAppearancePrefs auto = EquipmentStatusDashboardAppearancePrefs.defaults();
        EquipmentStatusDashboardAppearanceApplier.FlowLayoutSpec spec =
                EquipmentStatusDashboardAppearanceApplier.computeFlowSpec(auto, false, 0, 12, 12);
        Assertions.assertEquals(auto.cardWidth(), spec.wrapLength(), 0.01);
    }

    @Test
    void computeFlowSpec_fixedColumnsAddPaddingToTotalWidth() {
        EquipmentStatusDashboardAppearancePrefs fixed =
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
        EquipmentStatusDashboardAppearanceApplier.FlowLayoutSpec spec =
                EquipmentStatusDashboardAppearanceApplier.computeFlowSpec(fixed, false, 1200, 12, 12);
        Assertions.assertFalse(spec.fillViewport());
        Assertions.assertEquals(620, spec.wrapLength(), 0.01);
        Assertions.assertEquals(644, spec.totalWidth(), 0.01);
    }
}
