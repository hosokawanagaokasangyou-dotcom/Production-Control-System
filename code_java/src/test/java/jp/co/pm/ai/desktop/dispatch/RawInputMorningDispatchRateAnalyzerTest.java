package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class RawInputMorningDispatchRateAnalyzerTest {

    @TempDir Path tempDir;

    @Test
    void warnsWhenMorningRateLowDueToRawInputSameDay() throws Exception {
        String json =
                """
                {
                  "kwargs_packed": {
                    "timeline_events": [
                      {
                        "date": {"__t": "date", "v": "2026-06-10"},
                        "task_id": "",
                        "machine": "SEC+SEC機　湖南",
                        "machine_occupancy_key": "SEC機 湖南",
                        "start_dt": {"__t": "datetime", "v": "2026-06-10T08:45:00"},
                        "end_dt": {"__t": "datetime", "v": "2026-06-10T09:00:00"},
                        "event_kind": "machine_daily_startup"
                      },
                      {
                        "date": {"__t": "date", "v": "2026-06-10"},
                        "task_id": "",
                        "machine": "SL+スライス機1　湖南",
                        "machine_occupancy_key": "スライス機1 湖南",
                        "start_dt": {"__t": "datetime", "v": "2026-06-10T08:45:00"},
                        "end_dt": {"__t": "datetime", "v": "2026-06-10T09:00:00"},
                        "event_kind": "machine_daily_startup"
                      },
                      {
                        "date": {"__t": "date", "v": "2026-06-10"},
                        "task_id": "Y6-4",
                        "machine": "SEC+SEC機　湖南",
                        "machine_occupancy_key": "SEC機 湖南",
                        "start_dt": {"__t": "datetime", "v": "2026-06-10T13:05:00"},
                        "end_dt": {"__t": "datetime", "v": "2026-06-10T14:00:00"},
                        "event_kind": "machining"
                      },
                      {
                        "date": {"__t": "date", "v": "2026-06-10"},
                        "task_id": "Y5-1",
                        "machine": "SL+スライス機1　湖南",
                        "machine_occupancy_key": "スライス機1 湖南",
                        "start_dt": {"__t": "datetime", "v": "2026-06-10T09:00:00"},
                        "end_dt": {"__t": "datetime", "v": "2026-06-10T09:30:00"},
                        "event_kind": "machining"
                      }
                    ]
                  }
                }
                """;
        Path contract = tempDir.resolve("計画test設.json");
        Files.writeString(contract, json);

        LocalDate d = LocalDate.of(2026, 6, 10);
        var result =
                RawInputMorningDispatchRateAnalyzer.analyze(
                        contract, Map.of("Y6-4", d, "Y5-1", LocalDate.of(2026, 6, 9)));

        assertTrue(result.hasWarnings());
        assertEquals(1, result.lowRateDays().size());
        var day = result.lowRateDays().getFirst();
        assertEquals(d, day.date());
        assertTrue(day.morningRate() < RawInputMorningDispatchRateAnalyzer.RATE_THRESHOLD);
        assertEquals(1, day.rawInputSameDayTaskCount());
        assertEquals("Y6-4", day.rawInputSameDayTaskIds().getFirst());
    }

    @Test
    void noWarningWhenMorningRateAboveThreshold() throws Exception {
        String json =
                """
                {
                  "kwargs_packed": {
                    "timeline_events": [
                      {
                        "date": {"__t": "date", "v": "2026-06-10"},
                        "task_id": "",
                        "machine": "SEC+SEC機　湖南",
                        "machine_occupancy_key": "SEC機 湖南",
                        "start_dt": {"__t": "datetime", "v": "2026-06-10T08:45:00"},
                        "end_dt": {"__t": "datetime", "v": "2026-06-10T09:00:00"},
                        "event_kind": "machine_daily_startup"
                      },
                      {
                        "date": {"__t": "date", "v": "2026-06-10"},
                        "task_id": "Y6-4",
                        "machine": "SEC+SEC機　湖南",
                        "machine_occupancy_key": "SEC機 湖南",
                        "start_dt": {"__t": "datetime", "v": "2026-06-10T09:00:00"},
                        "end_dt": {"__t": "datetime", "v": "2026-06-10T12:30:00"},
                        "event_kind": "machining"
                      }
                    ]
                  }
                }
                """;
        Path contract = tempDir.resolve("計画test設.json");
        Files.writeString(contract, json);

        var result =
                RawInputMorningDispatchRateAnalyzer.analyze(
                        contract, Map.of("Y6-4", LocalDate.of(2026, 6, 10)));

        assertFalse(result.hasWarnings());
    }

    @Test
    void overlapMinutesClipsToMorningWindow() {
        LocalDate d = LocalDate.of(2026, 6, 10);
        long m =
                RawInputMorningDispatchRateAnalyzer.overlapMinutes(
                        java.time.LocalDateTime.of(d, java.time.LocalTime.of(9, 0)),
                        java.time.LocalDateTime.of(d, java.time.LocalTime.of(10, 0)),
                        d);
        assertEquals(60L, m);
        long zero =
                RawInputMorningDispatchRateAnalyzer.overlapMinutes(
                        java.time.LocalDateTime.of(d, java.time.LocalTime.of(14, 0)),
                        java.time.LocalDateTime.of(d, java.time.LocalTime.of(15, 0)),
                        d);
        assertEquals(0L, zero);
    }
}
