package jp.co.pm.ai.desktop.io;

import java.util.ArrayList;
import java.util.List;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

/** {@link TaskInputSourceRawGridIo#applyProcessingActualsDedupeByQuadKey}. */
class TaskInputSourceRawGridIoProcessingActualsTest {

    private static final String H_PROC = "工程名";
    private static final String H_MACH = "機械名";
    private static final String H_REQ = "依頼NO";
    private static final String H_DATE = "加工日";
    private static final String H_OTHER = "その他";

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
