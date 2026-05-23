package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Map;
import java.util.Optional;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.planning.stage2.core.Stage2PlanRowDispatchQtyMetrics;
import jp.co.pm.ai.planning.stage2.core.Stage2RollUnitLengthTables;

class Stage2InProgressNextDayRollInputTest {

    @Test
    void parseNonNegativeRollCountAllowsZeroAndIntegers() {
        assertEquals(Optional.of(0), Stage2InProgressNextDayRollInput.parseNonNegativeRollCount(""));
        assertEquals(Optional.of(3), Stage2InProgressNextDayRollInput.parseNonNegativeRollCount("3"));
        assertEquals(Optional.of(2), Stage2InProgressNextDayRollInput.parseNonNegativeRollCount("2.0"));
        assertTrue(Stage2InProgressNextDayRollInput.parseNonNegativeRollCount("-1").isEmpty());
        assertTrue(Stage2InProgressNextDayRollInput.parseNonNegativeRollCount("x").isEmpty());
    }

    @Test
    void resolveNextDayMetersUsesRollUnitLikeDispatchInteractive() {
        double unitM = 3045.0;
        assertEquals(
                0.0,
                Stage2InProgressNextDayRollInput.resolveNextDayMeters(0, 6090, unitM).orElse(-1.0),
                1e-9);
        assertEquals(
                6090.0,
                Stage2InProgressNextDayRollInput.resolveNextDayMeters(2, 6090, unitM).orElse(-1.0),
                1e-9);
        assertTrue(Stage2InProgressNextDayRollInput.resolveNextDayMeters(3, 6090, unitM).isEmpty());
    }

    @Test
    void validateRollInputRejectsExcessRolls() {
        Map<String, String> rowMap =
                Map.of(
                        "換算数量", "13530",
                        "実加工数", "2870",
                        "未加工", "10660",
                        "配台使用残数量", "10660",
                        "配台ロール数", "2",
                        "(原反)ロール単位長さ", "3045");
        Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM unitInfo =
                Stage2PlanRowDispatchQtyMetrics.dispatchSimulatorUnitMFromPlanRow(
                        rowMap, Stage2RollUnitLengthTables.empty());
        assertTrue(
                Stage2InProgressNextDayRollInput.validateRollInput("4", 10660, unitInfo).isPresent());
        assertTrue(
                Stage2InProgressNextDayRollInput.validateRollInput("3", 10660, unitInfo).isPresent());
        assertTrue(
                Stage2InProgressNextDayRollInput.validateRollInput("2", 10660, unitInfo).isEmpty());
        assertEquals(
                10660.0,
                Stage2InProgressNextDayRollInput.resolveNextDayMeters(2, 10660, unitInfo.unitM())
                        .orElse(-1.0),
                1e-9);
    }
}
