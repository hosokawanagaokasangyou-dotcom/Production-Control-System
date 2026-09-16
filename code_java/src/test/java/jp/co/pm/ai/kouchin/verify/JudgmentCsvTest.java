package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class JudgmentCsvTest {

    @TempDir Path tmp;

    @Test
    void missingManualIsEmpty() {
        JudgmentCsv.ManualLoad load =
                JudgmentCsv.loadManual(tmp.resolve("手動判定.csv"), "国分工場", new YearMonthKey(2026, 8));
        assertTrue(load.byKeiyaku().isEmpty());
        assertTrue(load.warnings().isEmpty());
    }

    @Test
    void manualAppliesOnlyMatchingMonthAndFactory() throws Exception {
        Path p = tmp.resolve("手動判定.csv");
        Files.writeString(
                p,
                "工場,対象月,契約NO,正とする側,理由\n"
                        + "国分工場,2026年8月度,189759V,①,1円差\n"
                        + "国分工場,2026年7月度,188000A,②,他月\n"
                        + "湖南工場,2026年8月度,189111B,②,他工場\n",
                StandardCharsets.UTF_8);
        JudgmentCsv.ManualLoad load = JudgmentCsv.loadManual(p, "国分工場", new YearMonthKey(2026, 8));
        assertTrue(load.warnings().isEmpty());
        assertEquals(1, load.byKeiyaku().size());
        assertEquals(1, load.byKeiyaku().get("189759V").side());
        assertEquals("1円差", load.byKeiyaku().get("189759V").reason());
    }

    @Test
    void missingPriorIsEmpty() {
        JudgmentCsv.PriorLoad load =
                JudgmentCsv.loadPrior(tmp.resolve("前月過不足.csv"), "国分工場", new YearMonthKey(2026, 8));
        assertTrue(load.rows().isEmpty());
        assertTrue(load.warnings().isEmpty());
    }

    @Test
    void priorParsesAmounts() throws Exception {
        Path p = tmp.resolve("前月過不足.csv");
        Files.writeString(
                p,
                "工場,対象月,契約NO,依頼NO,①東レ金額,②長岡金額,理由\n"
                        + "国分工場,2026年8月度,189000A,C8-1,0,1000,未反映\n",
                StandardCharsets.UTF_8);
        JudgmentCsv.PriorLoad load = JudgmentCsv.loadPrior(p, "国分工場", new YearMonthKey(2026, 8));
        assertEquals(1, load.rows().size());
        assertEquals(-1000.0, load.rows().get(0).diff(), 0.001);
    }
}
