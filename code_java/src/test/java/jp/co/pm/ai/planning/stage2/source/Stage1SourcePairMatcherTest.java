package jp.co.pm.ai.planning.stage2.source;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.reconciliation.KonanDailyReportLookup;

class Stage1SourcePairMatcherTest {

    @TempDir Path temp;

    @Test
    void matchPlanToDaily_picksNearestSameDayDaily() throws Exception {
        Path plan =
                temp.resolve("工程別加工計画_20260710_070500.xlsx");
        Path dailyNear = temp.resolve("加工日報発行問合せ_20260710_071800.csv");
        Path dailyFar = temp.resolve("加工日報発行問合せ_20260710_180000.csv");
        Files.writeString(plan, "plan", StandardCharsets.UTF_8);
        writeDailyCsv(dailyNear);
        writeDailyCsv(dailyFar);

        var planEntry =
                new NetworkSourceExtractionCatalog.SourceEntry(
                        plan,
                        LocalDateTime.of(2026, 7, 10, 7, 5),
                        NetworkSourceExtractionTimeSupport.SourceKind.FILENAME,
                        plan.getFileName().toString());
        var nearEntry =
                new NetworkSourceExtractionCatalog.SourceEntry(
                        dailyNear,
                        LocalDateTime.of(2026, 7, 10, 7, 18),
                        NetworkSourceExtractionTimeSupport.SourceKind.FILENAME,
                        dailyNear.getFileName().toString());
        var farEntry =
                new NetworkSourceExtractionCatalog.SourceEntry(
                        dailyFar,
                        LocalDateTime.of(2026, 7, 10, 18, 0),
                        NetworkSourceExtractionTimeSupport.SourceKind.FILENAME,
                        dailyFar.getFileName().toString());

        Stage1SourcePairMatcher.MatchedPair pair =
                Stage1SourcePairMatcher.matchPlanToDaily(
                        planEntry, List.of(farEntry, nearEntry));

        assertEquals(nearEntry, pair.dailyReport());
        assertEquals(13L, pair.deltaMinutes());
        assertFalse(pair.sameDayDailyMissing());
        assertFalse(pair.largeDeltaWarning());
    }

    @Test
    void matchPlanToDaily_flagsLargeDeltaWarning() {
        var planEntry =
                new NetworkSourceExtractionCatalog.SourceEntry(
                        temp.resolve("plan.xlsx"),
                        LocalDateTime.of(2026, 7, 10, 7, 5),
                        NetworkSourceExtractionTimeSupport.SourceKind.FILENAME,
                        "plan.xlsx");
        var dailyEntry =
                new NetworkSourceExtractionCatalog.SourceEntry(
                        temp.resolve("daily.csv"),
                        LocalDateTime.of(2026, 7, 10, 9, 30),
                        NetworkSourceExtractionTimeSupport.SourceKind.FILENAME,
                        "daily.csv");

        Stage1SourcePairMatcher.MatchedPair pair =
                Stage1SourcePairMatcher.matchPlanToDaily(planEntry, List.of(dailyEntry));

        assertTrue(pair.largeDeltaWarning());
    }

    @Test
    void matchPlanToDaily_sameDayMissing() {
        var planEntry =
                new NetworkSourceExtractionCatalog.SourceEntry(
                        temp.resolve("plan.xlsx"),
                        LocalDateTime.of(2026, 7, 10, 7, 5),
                        NetworkSourceExtractionTimeSupport.SourceKind.FILENAME,
                        "plan.xlsx");
        var otherDay =
                new NetworkSourceExtractionCatalog.SourceEntry(
                        temp.resolve("daily.csv"),
                        LocalDateTime.of(2026, 7, 9, 7, 18),
                        NetworkSourceExtractionTimeSupport.SourceKind.FILENAME,
                        "daily.csv");

        Stage1SourcePairMatcher.MatchedPair pair =
                Stage1SourcePairMatcher.matchPlanToDaily(planEntry, List.of(otherDay));

        assertTrue(pair.sameDayDailyMissing());
        assertNull(pair.dailyReport());
    }

    private static void writeDailyCsv(Path path) throws Exception {
        Files.writeString(
                path,
                "meta1\nmeta2\nmeta3\n依頼NO,工程名,機械名,加工日付,完了区分\n",
                StandardCharsets.UTF_8);
    }
}
