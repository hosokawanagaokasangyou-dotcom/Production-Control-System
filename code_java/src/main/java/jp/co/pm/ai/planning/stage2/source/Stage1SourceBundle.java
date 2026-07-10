package jp.co.pm.ai.planning.stage2.source;

import java.nio.file.Path;
import java.time.LocalDateTime;

/** 段階1成功時に固定する加工計画＋加工日報ソース束。 */
public record Stage1SourceBundle(
        LocalDateTime planExtractionTime,
        LocalDateTime dailyReportExtractionTime,
        long pairDeltaMinutes,
        String processingPlanPath,
        String dailyReportCsvPath,
        String dataExtractionWorkbookPath,
        long stage1CompletedAtEpochMillis) {

    public Stage1SourceBundle {
        processingPlanPath = normalizePathText(processingPlanPath);
        dailyReportCsvPath = normalizePathText(dailyReportCsvPath);
        dataExtractionWorkbookPath = normalizePathText(dataExtractionWorkbookPath);
    }

    public static Stage1SourceBundle fromMatchedPair(
            Stage1SourcePairMatcher.MatchedPair pair, long completedAtEpochMillis) {
        if (pair == null || pair.plan() == null || pair.dailyReport() == null) {
            throw new IllegalArgumentException("plan と dailyReport のペアが不完全です");
        }
        String planPath = pair.plan().path().toString();
        return new Stage1SourceBundle(
                pair.plan().extractionTime(),
                pair.dailyReport().extractionTime(),
                pair.deltaMinutes(),
                planPath,
                pair.dailyReport().path().toString(),
                planPath,
                completedAtEpochMillis);
    }

    public Path processingPlanPathNormalized() {
        return Path.of(processingPlanPath).toAbsolutePath().normalize();
    }

    public Path dailyReportCsvPathNormalized() {
        return Path.of(dailyReportCsvPath).toAbsolutePath().normalize();
    }

    private static String normalizePathText(String raw) {
        if (raw == null || raw.isBlank()) {
            return "";
        }
        return Path.of(raw.strip()).toAbsolutePath().normalize().toString();
    }
}
