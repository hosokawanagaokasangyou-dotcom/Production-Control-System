package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class InspectionSheetIndexProgressTest {

    @Test
    void format_walk_showsFoundCountWithoutTotal() {
        assertEquals(
                "検査表索引 フォルダ走査中 0 件",
                InspectionSheetIndexProgress.format(
                        InspectionSheetIndexProgress.PHASE_WALK, 0, 0));
        assertEquals(
                "検査表索引 フォルダ走査中 12 件",
                InspectionSheetIndexProgress.format(
                        InspectionSheetIndexProgress.PHASE_WALK, 12, 0));
    }

    @Test
    void format_index_showsDoneOverTotal() {
        assertEquals(
                "検査表索引 索引更新中 3 / 10",
                InspectionSheetIndexProgress.format(
                        InspectionSheetIndexProgress.PHASE_INDEX, 3, 10));
    }

    @Test
    void fraction_unknownTotal_isNaN() {
        assertTrue(Double.isNaN(InspectionSheetIndexProgress.fraction(12, 0)));
        assertTrue(Double.isNaN(InspectionSheetIndexProgress.fraction(0, -1)));
    }

    @Test
    void fraction_knownTotal_isRatio() {
        assertEquals(0.3, InspectionSheetIndexProgress.fraction(3, 10), 1e-9);
        assertEquals(1.0, InspectionSheetIndexProgress.fraction(10, 10), 1e-9);
    }

    @Test
    void shouldPublishUi_throttlesUntilIntervalOrForce() {
        long last = 1_000_000_000L;
        assertFalse(
                InspectionSheetIndexProgress.shouldPublishUi(
                        last, last + InspectionSheetIndexProgress.UI_THROTTLE_NS - 1, false));
        assertTrue(
                InspectionSheetIndexProgress.shouldPublishUi(
                        last, last + InspectionSheetIndexProgress.UI_THROTTLE_NS, false));
        assertTrue(
                InspectionSheetIndexProgress.shouldPublishUi(
                        last, last + 1, true));
    }
}
