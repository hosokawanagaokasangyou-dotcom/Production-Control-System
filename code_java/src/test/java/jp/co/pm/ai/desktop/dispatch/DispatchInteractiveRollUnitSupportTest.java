package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Map;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.planning.stage2.core.Stage2PlanRowDispatchQtyMetrics;
import jp.co.pm.ai.planning.stage2.core.Stage2RollUnitLengthTables;

class DispatchInteractiveRollUnitSupportTest {

    @Test
    void rollAlignmentAndLargestMultiple() {
        assertTrue(Stage2PlanRowDispatchQtyMetrics.isQtyAlignedToRollUnit(6000, 3000));
        assertFalse(Stage2PlanRowDispatchQtyMetrics.isQtyAlignedToRollUnit(3001, 3000));
        assertEquals(
                6000.0,
                Stage2PlanRowDispatchQtyMetrics.largestRollMultipleNotExceeding(6090, 3000),
                1e-9);
    }

    @Test
    void moveRollCountAndMetersPreview() {
        assertEquals(29, DispatchInteractiveRollUnitSupport.maxMoveRollCount(6090, 210));
        assertEquals(29, DispatchInteractiveRollUnitSupport.defaultMoveRollCount(6090, 210));
        assertEquals(6090.0, DispatchInteractiveRollUnitSupport.metersForRollCount(29, 210), 1e-9);
        assertEquals(420.0, DispatchInteractiveRollUnitSupport.metersForRollCount(2, 210), 1e-9);
        String preview = DispatchInteractiveRollUnitSupport.formatMoveMetersPreview(29, 210);
        assertTrue(preview.contains("6090"));
        assertTrue(preview.contains("29"));
    }

    @Test
    void unitMFromDispatchRollColumns() {
        Map<String, String> row =
                Map.of(
                        "換算数量", "6090",
                        "実加工数", "0",
                        "未加工", "6090",
                        "配台使用残数量", "6090",
                        "配台ロール数", "2",
                        "(原反)ロール単位長さ", "3045");
        Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM u =
                Stage2PlanRowDispatchQtyMetrics.dispatchSimulatorUnitMFromPlanRow(
                        row, Stage2RollUnitLengthTables.empty());
        assertTrue(u.fromDispatchRollColumns());
        assertEquals(3045.0, u.unitM(), 1e-6);
        assertEquals(2.0, u.dispatchRollCount(), 1e-9);
    }
}
