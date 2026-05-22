package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.planning.stage2.core.Stage2PlanRowDispatchQtyMetrics;
import jp.co.pm.ai.planning.stage2.core.Stage2RollUnitLengthTables;

/** 配台ロール数列採用後の (原反)ロール単位長さ・黄強調の整合。 */
final class SpreadsheetTabularSupportPlanInputYellowTest {

    @Test
    void row950mRoll100_doesNotNeedEffectiveRollYellow() {
        List<String> headers =
                List.of("換算数量", "未加工", "実加工数", "(原反)ロール単位長さ");
        List<String> cells = List.of("950", "0", "0", "100");
        Map<String, String> row =
                Map.of(
                        "換算数量", "950",
                        "未加工", "0",
                        "実加工数", "0",
                        "(原反)ロール単位長さ", "100");
        assertTrue(
                Stage2PlanRowDispatchQtyMetrics.stage2SimulatorUsesDispatchRollCountColumns(
                        row, Stage2RollUnitLengthTables.empty()));
        assertFalse(
                SpreadsheetTabularSupport.planInputRawRollUnitCellYellowHighlightForTest(
                        headers, cells, Stage2RollUnitLengthTables.empty()));
    }

    @Test
    void derivedRollColumns800over95_noYellow() {
        List<String> headers = List.of("換算数量", "未加工", "(原反)ロール単位長さ");
        List<String> cells = List.of("800", "800", "95");
        Map<String, String> row =
                Map.of("換算数量", "800", "未加工", "800", "(原反)ロール単位長さ", "95");
        assertTrue(
                Stage2PlanRowDispatchQtyMetrics.stage2SimulatorUsesDispatchRollCountColumns(
                        row, Stage2RollUnitLengthTables.empty()));
        assertFalse(
                SpreadsheetTabularSupport.planInputRawRollUnitCellYellowHighlightForTest(
                        headers, cells, Stage2RollUnitLengthTables.empty()));
    }

    @Test
    void legacyEffectiveRollYellow_whenRollCountZero() {
        List<String> headers =
                List.of("換算数量", "未加工", "配台使用残数量", "配台ロール数", "(原反)ロール単位長さ");
        List<String> cells = List.of("800", "800", "800", "0", "95");
        Map<String, String> row =
                Map.of(
                        "換算数量", "800",
                        "未加工", "800",
                        "配台使用残数量", "800",
                        "配台ロール数", "0",
                        "(原反)ロール単位長さ", "95");
        assertFalse(
                Stage2PlanRowDispatchQtyMetrics.stage2SimulatorUsesDispatchRollCountColumns(
                        row, Stage2RollUnitLengthTables.empty()));
        assertTrue(
                SpreadsheetTabularSupport.planInputRawRollUnitCellYellowHighlightForTest(
                        headers, cells, Stage2RollUnitLengthTables.empty()));
    }
}
