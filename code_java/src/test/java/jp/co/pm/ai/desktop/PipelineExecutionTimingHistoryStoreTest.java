package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.Test;

class PipelineExecutionTimingHistoryStoreTest {

    @Test
    void computeStatsSummarizesDurations() {
        List<PipelineExecutionTimingSample> samples =
                List.of(
                        new PipelineExecutionTimingSample(
                                PipelineExecutionTimingKind.STAGE1, 1L, 10_000L),
                        new PipelineExecutionTimingSample(
                                PipelineExecutionTimingKind.STAGE1, 2L, 20_000L),
                        new PipelineExecutionTimingSample(
                                PipelineExecutionTimingKind.STAGE1, 3L, 30_000L));
        PipelineExecutionTimingHistoryStore.Stats stats =
                PipelineExecutionTimingHistoryStore.computeStats(samples);
        assertEquals(3L, stats.count());
        assertEquals(20.0, stats.avgSec(), 0.01);
        assertEquals(20.0, stats.medianSec(), 0.01);
        assertEquals(10.0, stats.minSec(), 0.01);
        assertEquals(30.0, stats.maxSec(), 0.01);
    }

    @Test
    void computeHistogramBinsCoverAllSamples() {
        List<PipelineExecutionTimingSample> samples =
                List.of(
                        new PipelineExecutionTimingSample(
                                PipelineExecutionTimingKind.STAGE2_0, 1L, 1_000L, "pc-a", "10.0.0.1"),
                        new PipelineExecutionTimingSample(
                                PipelineExecutionTimingKind.STAGE2_0, 2L, 2_000L, "pc-a", "10.0.0.1"),
                        new PipelineExecutionTimingSample(
                                PipelineExecutionTimingKind.STAGE2_0, 3L, 9_000L, "pc-b", "10.0.0.2"));
        List<PipelineExecutionTimingHistoryStore.HistogramBin> bins =
                PipelineExecutionTimingHistoryStore.computeHistogram(samples, 3);
        assertEquals(3, bins.size());
        int total = bins.stream().mapToInt(PipelineExecutionTimingHistoryStore.HistogramBin::count).sum();
        assertEquals(3, total);
        assertTrue(bins.stream().allMatch(b -> b.count() >= 0));
    }

    @Test
    void parseKindMigratesLegacyStage2ToStage20() throws Exception {
        var parseKind =
                PipelineExecutionTimingHistoryStore.class.getDeclaredMethod("parseKind", String.class);
        parseKind.setAccessible(true);
        assertEquals(
                PipelineExecutionTimingKind.STAGE2_0,
                parseKind.invoke(null, "STAGE2"));
        assertEquals(
                PipelineExecutionTimingKind.STAGE2_0,
                parseKind.invoke(null, "STAGE2_0"));
    }

    @Test
    void mergeSamplesUnionsDiskAndMemoryWithoutDuplicates() {
        List<PipelineExecutionTimingSample> disk =
                List.of(
                        new PipelineExecutionTimingSample(
                                PipelineExecutionTimingKind.STAGE1, 100L, 5_000L, "remote", "192.168.0.2"));
        List<PipelineExecutionTimingSample> memory =
                List.of(
                        new PipelineExecutionTimingSample(
                                PipelineExecutionTimingKind.STAGE1, 200L, 6_000L, "local", "192.168.0.1"),
                        new PipelineExecutionTimingSample(
                                PipelineExecutionTimingKind.STAGE1, 100L, 5_000L, "remote", "192.168.0.2"));
        List<PipelineExecutionTimingSample> merged =
                PipelineExecutionTimingHistoryStore.mergeSamples(disk, memory);
        assertEquals(2, merged.size());
        assertEquals(100L, merged.get(0).finishedAtEpochMs());
        assertEquals(200L, merged.get(1).finishedAtEpochMs());
    }
}
