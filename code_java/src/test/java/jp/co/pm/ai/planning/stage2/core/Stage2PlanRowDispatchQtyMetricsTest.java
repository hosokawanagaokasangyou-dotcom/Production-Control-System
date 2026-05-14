package jp.co.pm.ai.planning.stage2.core;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class Stage2PlanRowDispatchQtyMetricsTest {

    @Test
    void unprocessedPositive_splitsRemainingAndDone() {
        Map<String, String> row = baseRow();
        row.put("換算数量", "100");
        row.put("未加工", "40");
        Optional<Stage2PlanRowDispatchQtyMetrics.Metrics> m =
                Stage2PlanRowDispatchQtyMetrics.compute(row, Stage2RollUnitLengthTables.empty());
        assertTrue(m.isPresent());
        assertEquals(40.0, m.get().remainingM(), 1e-9);
        assertEquals(60.0, m.get().doneM(), 1e-9);
        assertEquals(100.0, m.get().qtyTotalForDispatchM(), 1e-9);
    }

    @Test
    void allUnprocessedWhenZeroUnprocessedAndZeroActual() {
        Map<String, String> row = baseRow();
        row.put("換算数量", "100");
        row.put("未加工", "0");
        row.put("実加工数", "0");
        Optional<Stage2PlanRowDispatchQtyMetrics.Metrics> m =
                Stage2PlanRowDispatchQtyMetrics.compute(row, Stage2RollUnitLengthTables.empty());
        assertTrue(m.isPresent());
        assertEquals(100.0, m.get().remainingM(), 1e-9);
        assertEquals(0.0, m.get().doneM(), 1e-9);
    }

    @Test
    void missingUnprocessedColumn_returnsEmpty() {
        Map<String, String> row = new HashMap<>();
        row.put("換算数量", "100");
        assertTrue(
                Stage2PlanRowDispatchQtyMetrics.compute(row, Stage2RollUnitLengthTables.empty())
                        .isEmpty());
    }

    @Test
    void rollUnitFromUsedRawTable(@TempDir Path tmp) throws Exception {
        Path code = tmp.resolve("code");
        Files.createDirectories(code);
        Files.writeString(
                code.resolve("使用原反,ロール単位の長さ.txt"),
                "使用原反,ロール単位の長さ\nMY-RAW-KEY,250\n",
                StandardCharsets.UTF_8);
        Files.writeString(
                code.resolve("製品名,ロール単位の長さ.txt"),
                "製品名,ロール単位の長さ\n",
                StandardCharsets.UTF_8);
        Stage2RollUnitLengthTables tables = Stage2RollUnitLengthTables.load(tmp);
        Map<String, String> row = baseRow();
        row.put("換算数量", "125");
        row.put("未加工", "0");
        row.put("実加工数", "10");
        row.put("使用原反", "MY-RAW-KEY");
        Optional<Stage2PlanRowDispatchQtyMetrics.Metrics> m = Stage2PlanRowDispatchQtyMetrics.compute(row, tables);
        assertTrue(m.isPresent());
        assertEquals(250.0, m.get().remainingM(), 1e-9);
    }

    private static Map<String, String> baseRow() {
        return new HashMap<>();
    }
}
