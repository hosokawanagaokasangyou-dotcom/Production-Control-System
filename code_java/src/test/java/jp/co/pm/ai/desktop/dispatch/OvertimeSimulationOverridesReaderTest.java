package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class OvertimeSimulationOverridesReaderTest {

    @TempDir Path tempDir;

    @Test
    void summarize_countsWorkingAndOvertimeCells() throws Exception {
        Path json = tempDir.resolve("overtime_simulation_overrides.json");
        Files.writeString(
                json,
                """
                {
                  "working_overrides": {
                    "2026-05-28": { "A": true, "B": false },
                    "2026-05-29": { "C": true }
                  },
                  "overtime_minutes": {
                    "2026-05-28": { "A": 60, "D": 30 }
                  }
                }
                """,
                StandardCharsets.UTF_8);

        Stage21TrialSnapshotStore.OverrideSummary s =
                OvertimeSimulationOverridesReader.summarize(json);
        assertEquals(2, s.workOn());
        assertEquals(1, s.workOff());
        assertEquals(2, s.overtimeCells());
        assertTrue(s.formatSummaryLine().contains("休日出勤（○化）: 2"));
    }
}
