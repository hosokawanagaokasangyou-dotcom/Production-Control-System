package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class MachineCalendarBlockIndexTest {

    @Test
    void loadOutcomeFromJsonFile_parsesBlocks(@TempDir Path tmp) throws Exception {
        Path f = tmp.resolve("c.json");
        Files.writeString(
                f,
                "{\"blocks\": {\"M1\": [\"2026-05-03T00:00:00\"], \"P\\tQ\": [\"2026-05-04\"]}}",
                StandardCharsets.UTF_8);
        MachineCalendarBlockIndex.LoadOutcome lo = MachineCalendarBlockIndex.loadOutcomeFromJsonFile(f);
        assertTrue(!lo.index().isEmpty());
        assertTrue(
                lo.index()
                        .isBlockedDay("P", "Q", LocalDate.of(2026, 5, 4)));
        assertTrue(
                lo.index()
                        .isBlockedDay("", "M1", LocalDate.of(2026, 5, 3)));
    }

    @Test
    void loadOutcomeFromJsonFile_missingFile(@TempDir Path tmp) throws Exception {
        Path f = tmp.resolve("nope.json");
        MachineCalendarBlockIndex.LoadOutcome lo = MachineCalendarBlockIndex.loadOutcomeFromJsonFile(f);
        assertTrue(lo.index().isEmpty());
        assertEquals("missing_file", lo.pythonJsonError());
    }

    @Test
    void loadOutcomeFromJsonFile_prettyPrinted(@TempDir Path tmp) throws Exception {
        Path out = tmp.resolve("machine_calendar_blocks.json");
        Files.writeString(
                out,
                "{\n  \"blocks\" : {\n    \"A\" : [ \"2026-01-01\" ]\n  }\n}\n",
                StandardCharsets.UTF_8);
        MachineCalendarBlockIndex.LoadOutcome lo =
                MachineCalendarBlockIndex.loadOutcomeFromJsonFile(out);
        assertTrue(lo.index().isBlockedDay("", "A", LocalDate.of(2026, 1, 1)));
    }

    @Test
    void matchesEquipmentKey_variants() {
        assertTrue(MachineCalendarBlockIndex.matchesEquipmentKey("M", "P", "M"));
        assertTrue(MachineCalendarBlockIndex.matchesEquipmentKey("P\tM", "P", "M"));
    }
}
