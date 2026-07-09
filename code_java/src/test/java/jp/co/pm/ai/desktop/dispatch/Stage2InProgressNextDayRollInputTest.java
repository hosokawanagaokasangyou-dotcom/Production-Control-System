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

    @Test
    void maxRollsForCapUsesMinOfAladdinTodayAndRemaining() {
        assertEquals(2, Stage2InProgressNextDayRollInput.maxRollsForCap(6090, 10660, 3045));
        assertEquals(1, Stage2InProgressNextDayRollInput.maxRollsForCap(4000, 10660, 3045));
        assertEquals(0, Stage2InProgressNextDayRollInput.maxRollsForCap(3000, 10660, 3045));
    }

    @Test
    void nextDayTargetAssumesAladdinTodayCompletes() {
        assertEquals(
                3000.0,
                Stage2InProgressNextDayRollInput.nextDayTargetMetersAssumingAladdinTodayComplete(
                        5700, 300, 3000),
                1e-9);
        assertEquals(
                600.0,
                Stage2InProgressNextDayRollInput.nextDayTargetMetersAssumingAladdinTodayComplete(
                        600, 900, 0),
                1e-9);
        assertEquals(
                10,
                Stage2InProgressNextDayRollInput.defaultRollCountAssumingAladdinTodayComplete(
                        5700, 300, 3000, 300));
        assertEquals(
                2,
                Stage2InProgressNextDayRollInput.defaultRollCountAssumingAladdinTodayComplete(
                        600, 900, 0, 300));
    }

    @Test
    void validateExcludeRollInputCapsOnlyByRemainingNotAladdinToday() {
        Map<String, String> rowMap =
                Map.of(
                        "換算数量", "8000",
                        "実加工数", "0",
                        "未加工", "8000",
                        "配台使用残数量", "8000",
                        "(原反)ロール単位長さ", "3045");
        Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM unitInfo =
                Stage2PlanRowDispatchQtyMetrics.dispatchSimulatorUnitMFromPlanRow(
                        rowMap, Stage2RollUnitLengthTables.empty());
        // アラジン当日量(6090m=2ロール)を超える3ロール(9135m)は残量(8000m)超過のため拒否。
        assertTrue(
                Stage2InProgressNextDayRollInput.validateExcludeRollInput("3", 8000, unitInfo)
                        .isPresent());
        // 残量(8000m)に収まる2ロール(6090m)はアラジン当日量と一致するため許可（従来どおり）。
        assertTrue(
                Stage2InProgressNextDayRollInput.validateExcludeRollInput("2", 8000, unitInfo)
                        .isEmpty());
    }

    @Test
    void validateExcludeRollInputAllowsExcludingFullRemainingBeyondAladdinToday() {
        Map<String, String> rowMap =
                Map.of(
                        "換算数量", "10000",
                        "実加工数", "0",
                        "未加工", "10000",
                        "配台使用残数量", "10000",
                        "(原反)ロール単位長さ", "200");
        Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM unitInfo =
                Stage2PlanRowDispatchQtyMetrics.dispatchSimulatorUnitMFromPlanRow(
                        rowMap, Stage2RollUnitLengthTables.empty());
        // アラジン当日量は5000m(25ロール)だが、上限は残量(10000m=50ロール)まで。
        assertTrue(
                Stage2InProgressNextDayRollInput.validateExcludeRollInput("50", 10000, unitInfo)
                        .isEmpty());
        assertTrue(
                Stage2InProgressNextDayRollInput.validateExcludeRollInput("51", 10000, unitInfo)
                        .isPresent());
    }
}
