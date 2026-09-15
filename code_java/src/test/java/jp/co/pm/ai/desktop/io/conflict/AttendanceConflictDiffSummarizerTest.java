package jp.co.pm.ai.desktop.io.conflict;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

class AttendanceConflictDiffSummarizerTest {

    @Test
    void summarizesMemberAddedAndCellChanges() {
        Path json = Path.of("attendance-data.json").toAbsolutePath().normalize();
        Path xlsx = Path.of("勤怠・機械カレンダー.xlsx").toAbsolutePath().normalize();
        String base =
                """
                {"member_roster":[{"name":"佐藤"}],"member_attendance":{"2026-09-01":{"佐藤":{"day_preset":"出勤"}}},"company_calendar":{"days":{}}}
                """;
        String disk =
                """
                {"member_roster":[{"name":"佐藤"},{"name":"山田"}],"member_attendance":{"2026-09-01":{"佐藤":{"day_preset":"公休"},"山田":{"day_preset":"出勤"}}},"company_calendar":{"days":{}}}
                """;
        byte[] xBase = new byte[] {1};
        byte[] xDisk = new byte[] {2};
        String summary =
                new AttendanceConflictDiffSummarizer()
                        .summarize(
                                Map.of(json, base.getBytes(StandardCharsets.UTF_8), xlsx, xBase),
                                Map.of(json, disk.getBytes(StandardCharsets.UTF_8), xlsx, xDisk),
                                List.of(json, xlsx));
        assertTrue(summary.contains("山田"), summary);
        assertTrue(summary.contains("追加"), summary);
        assertTrue(
                summary.contains("Excel") || summary.contains("xlsx") || summary.contains("カレンダー"),
                summary);
        assertTrue(summary.contains("セル"), summary);
    }

    @Test
    void fallbackWhenJsonBroken() {
        Path json = Path.of("attendance-data.json").toAbsolutePath().normalize();
        String summary =
                new AttendanceConflictDiffSummarizer()
                        .summarize(
                                Map.of(json, "{".getBytes(StandardCharsets.UTF_8)),
                                Map.of(json, "}".getBytes(StandardCharsets.UTF_8)),
                                List.of(json));
        assertTrue(
                summary.contains("詳細差分を生成できませんでした") || summary.contains("attendance-data"),
                summary);
    }

    @Test
    void choiceEnumStable() {
        assertEquals(3, ConflictSaveChoice.values().length);
    }

    @Test
    void summarizesInactiveFromChangeWithoutTreatingAsDelete() {
        Path json = Path.of("attendance-data.json").toAbsolutePath().normalize();
        String base =
                """
                {"member_roster":[{"name":"菅沼　めぐみ","primary_role":"後加工"}],"member_attendance":{},"company_calendar":{"days":{}}}
                """;
        String disk =
                """
                {"member_roster":[{"name":"菅沼　めぐみ","primary_role":"後加工","inactive_from":"2026-09-15"}],"member_attendance":{},"company_calendar":{"days":{}}}
                """;
        String summary =
                new AttendanceConflictDiffSummarizer()
                        .summarize(
                                Map.of(json, base.getBytes(StandardCharsets.UTF_8)),
                                Map.of(json, disk.getBytes(StandardCharsets.UTF_8)),
                                List.of(json));
        assertTrue(summary.contains("菅沼"), summary);
        assertTrue(summary.contains("異動") || summary.contains("inactive"), summary);
        assertTrue(!summary.contains("削除"), summary);
    }

    @Test
    void summarizesReturnedOnChange() {
        Path json = Path.of("attendance-data.json").toAbsolutePath().normalize();
        String base =
                """
                {"member_roster":[{"name":"菅沼　めぐみ","primary_role":"後加工","inactive_from":"2026-09-15"}],"member_attendance":{},"company_calendar":{"days":{}}}
                """;
        String disk =
                """
                {"member_roster":[{"name":"菅沼　めぐみ","primary_role":"後加工","inactive_from":"2026-09-15","returned_on":"2026-11-01"}],"member_attendance":{},"company_calendar":{"days":{}}}
                """;
        String summary =
                new AttendanceConflictDiffSummarizer()
                        .summarize(
                                Map.of(json, base.getBytes(StandardCharsets.UTF_8)),
                                Map.of(json, disk.getBytes(StandardCharsets.UTF_8)),
                                List.of(json));
        assertTrue(summary.contains("復帰"), summary);
        assertTrue(summary.contains("2026-11-01"), summary);
    }
}
