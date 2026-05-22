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
                                PipelineExecutionTimingKind.STAGE2, 1L, 1_000L),
                        new PipelineExecutionTimingSample(
                                PipelineExecutionTimingKind.STAGE2, 2L, 2_000L),
                        new PipelineExecutionTimingSample(
                                PipelineExecutionTimingKind.STAGE2, 3L, 9_000L));
        List<PipelineExecutionTimingHistoryStore.HistogramBin> bins =
                PipelineExecutionTimingHistoryStore.computeHistogram(samples, 3);
        assertEquals(3, bins.size());
        int total = bins.stream().mapToInt(PipelineExecutionTimingHistoryStore.HistogramBin::count).sum();
        assertEquals(3, total);
        assertTrue(bins.stream().allMatch(b -> b.count() >= 0));
    }
}
