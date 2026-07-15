package jp.co.pm.ai.planning.stage2.source;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class Stage1SourceBundleIoTest {
    @TempDir Path temp;

    @Test
    void readRejectsMissingRequiredValuesWithReason() throws Exception {
        Path path = temp.resolve("broken.json");
        Files.writeString(path, "{}");
        IOException ex = assertThrows(IOException.class, () -> Stage1SourceBundleIo.read(path));
        assertTrue(ex.getMessage().contains("planExtractionTime"));
    }

    @Test
    void readRejectsDecimalRequiredLong() throws Exception {
        Path path = temp.resolve("decimal-long.json");
        Files.writeString(
                path,
                validJson().replace("\"pairDeltaMinutes\": 13", "\"pairDeltaMinutes\": 13.5"));

        IOException ex = assertThrows(IOException.class, () -> Stage1SourceBundleIo.read(path));
        assertTrue(ex.getMessage().contains("pairDeltaMinutes"));
    }

    @Test
    void readRejectsNumericRequiredText() throws Exception {
        Path path = temp.resolve("numeric-text.json");
        Files.writeString(
                path,
                validJson().replace("\"processingPlanPath\": \"plan.xlsx\"", "\"processingPlanPath\": 123"));

        IOException ex = assertThrows(IOException.class, () -> Stage1SourceBundleIo.read(path));
        assertTrue(ex.getMessage().contains("processingPlanPath"));
    }

    @Test
    void readRejectsBooleanRequiredText() throws Exception {
        Path path = temp.resolve("boolean-text.json");
        Files.writeString(
                path,
                validJson().replace("\"dailyReportCsvPath\": \"daily.csv\"", "\"dailyReportCsvPath\": true"));

        IOException ex = assertThrows(IOException.class, () -> Stage1SourceBundleIo.read(path));
        assertTrue(ex.getMessage().contains("dailyReportCsvPath"));
    }

    @Test
    void moveFailureRemovesOldAndTemporaryJson() throws Exception {
        Path path = temp.resolve("bundle.json");
        Files.writeString(path, "{}");
        assertThrows(IOException.class, () -> Stage1SourceBundleIo.writeWithMove(path, validBundle(), (tmp, target) -> { throw new IOException("move"); }));
        assertFalse(Files.exists(path));
        try (var files = Files.list(temp)) {
            assertFalse(files.anyMatch(p -> p.getFileName().toString().endsWith(".tmp")));
        }
    }

    private Stage1SourceBundle validBundle() {
        return new Stage1SourceBundle(
                LocalDateTime.of(2026, 7, 10, 7, 5),
                LocalDateTime.of(2026, 7, 10, 7, 18),
                13L,
                temp.resolve("plan.xlsx").toString(),
                temp.resolve("daily.csv").toString(),
                temp.resolve("plan.xlsx").toString(),
                1L);
    }

    private static String validJson() {
        return """
                {
                  "planExtractionTime": "2026-07-10T07:05:00",
                  "dailyReportExtractionTime": "2026-07-10T07:18:00",
                  "pairDeltaMinutes": 13,
                  "processingPlanPath": "plan.xlsx",
                  "dailyReportCsvPath": "daily.csv",
                  "dataExtractionWorkbookPath": "plan.xlsx",
                  "stage1CompletedAtEpochMillis": 1
                }
                """;
    }
}
