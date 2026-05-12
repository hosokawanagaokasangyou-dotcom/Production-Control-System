package jp.co.pm.ai.desktop.io;

import java.util.ArrayList;
import java.util.List;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

/**
 * {@link TaskInputSourceRawGridIo#applyProcessingActualsDedupeByQuadKey} and {@link
 * TaskInputSourceRawGridIo#applyProcessingActualsDisplaySteps}.
 */
class TaskInputSourceRawGridIoProcessingActualsTest {

    private static final String H_PROC = "工程名";
    private static final String H_MACH = "機械名";
    private static final String H_REQ = "依頼NO";
    private static final String H_DATE = "加工日";
    private static final String H_OTHER = "その他";

    @Test
    void displaySteps_trimsRowsAboveInspectionNoMarker() {
        List<List<String>> data = new ArrayList<>();
        data.add(List.of("倉庫", "520201"));
        data.add(List.of("加工日付", "2026年04月27日"));
        data.add(List.of("検査NO", "工程名", "機械名"));
        data.add(List.of("1", "P1", "M1"));

        PlanInputTabularIo.TabularSheet raw =
                new PlanInputTabularIo.TabularSheet(List.of("列1", "列2"), data);
        PlanInputTabularIo.TabularSheet out =
                TaskInputSourceRawGridIo.applyProcessingActualsDisplaySteps(raw);

        Assertions.assertEquals(List.of("検査NO", "工程名", "機械名"), out.headers());
        Assertions.assertEquals(1, out.rows().size());
        Assertions.assertEquals(List.of("1", "P1", "M1"), out.rows().get(0));
    }

    @Test
    void displaySteps_acceptsFullwidthInspectionNoMarker() {
        List<List<String>> data =
                List.of(
                        List.of("meta"),
                        List.of("検査ＮＯ", "工程名"),
                        List.of("x", "y"));
        PlanInputTabularIo.TabularSheet raw =
                new PlanInputTabularIo.TabularSheet(List.of("列1"), data);
        PlanInputTabularIo.TabularSheet out =
                TaskInputSourceRawGridIo.applyProcessingActualsDisplaySteps(raw);

        Assertions.assertEquals(List.of("検査ＮＯ", "工程名"), out.headers());
        Assertions.assertEquals(List.of("x", "y"), out.rows().get(0));
    }

    @Test
    void displaySteps_fallbackFirstFourRowsWhenMarkerAbsent() {
        List<List<String>> data = new ArrayList<>();
        for (int i = 0; i < 6; i++) {
            data.add(List.of("r" + i, "c" + i));
        }
        PlanInputTabularIo.TabularSheet raw =
                new PlanInputTabularIo.TabularSheet(List.of("列1"), data);
        PlanInputTabularIo.TabularSheet out =
                TaskInputSourceRawGridIo.applyProcessingActualsDisplaySteps(raw);

        Assertions.assertEquals(List.of("r4", "c4"), out.headers());
        Assertions.assertEquals(1, out.rows().size());
        Assertions.assertEquals(List.of("r5", "c5"), out.rows().get(0));
    }

    @Test
    void dedupe_keepsFirstAmongDuplicates() {
        List<String> headers = List.of(H_PROC, H_MACH, H_REQ, H_DATE, H_OTHER);
        List<List<String>> rows = new ArrayList<>();
        rows.add(List.of("P1", "M1", "R1", "2025/1/1", "a"));
        rows.add(List.of("P1", "M1", "R1", "2025/1/1", "drop"));
        rows.add(List.of("P2", "M2", "R2", "2025/1/2", "keep"));

        PlanInputTabularIo.TabularSheet in = new PlanInputTabularIo.TabularSheet(headers, rows);
        PlanInputTabularIo.TabularSheet out =
                TaskInputSourceRawGridIo.applyProcessingActualsDedupeByQuadKey(in);

        Assertions.assertEquals(2, out.rows().size());
        Assertions.assertEquals("a", out.rows().get(0).get(4));
        Assertions.assertEquals("keep", out.rows().get(1).get(4));
    }

    @Test
    void dedupe_noopWhenMissingColumn() {
        List<String> headers = List.of(H_PROC, H_MACH, H_DATE);
        List<List<String>> rows =
                List.of(
                        List.of("P1", "M1", "2025/1/1"),
                        List.of("P1", "M1", "2025/1/1"));
        PlanInputTabularIo.TabularSheet in = new PlanInputTabularIo.TabularSheet(headers, rows);
        PlanInputTabularIo.TabularSheet out =
                TaskInputSourceRawGridIo.applyProcessingActualsDedupeByQuadKey(in);

        Assertions.assertEquals(2, out.rows().size());
    }

    @Test
    void dedupe_acceptsFullwidthIraiHeaderAlias() {
        List<String> headers = List.of(H_PROC, H_MACH, "依頼ＮＯ", H_DATE);
        List<List<String>> rows = new ArrayList<>();
        rows.add(List.of("P1", "M1", "R1", "2025/1/1"));
        rows.add(List.of("P1", "M1", "R1", "2025/1/1"));

        PlanInputTabularIo.TabularSheet in = new PlanInputTabularIo.TabularSheet(headers, rows);
        PlanInputTabularIo.TabularSheet out =
                TaskInputSourceRawGridIo.applyProcessingActualsDedupeByQuadKey(in);

        Assertions.assertEquals(1, out.rows().size());
    }
}
