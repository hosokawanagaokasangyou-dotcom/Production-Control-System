package jp.co.pm.ai.planning.stage2.source;

import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.Duration;
import java.util.Optional;

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

    public Optional<String> validationError() {
        if (planExtractionTime == null) return Optional.of("planExtractionTime がありません");
        if (dailyReportExtractionTime == null) return Optional.of("dailyReportExtractionTime がありません");
        long expected = Math.abs(Duration.between(planExtractionTime, dailyReportExtractionTime).toMinutes());
        if (pairDeltaMinutes < 0 || pairDeltaMinutes != expected) return Optional.of("pairDeltaMinutes が不正です");
        Optional<String> plan = validatePath("processingPlanPath", processingPlanPath);
        if (plan.isPresent()) return plan;
        Optional<String> daily = validatePath("dailyReportCsvPath", dailyReportCsvPath);
        if (daily.isPresent()) return daily;
        Optional<String> extraction = validatePath("dataExtractionWorkbookPath", dataExtractionWorkbookPath);
        if (extraction.isPresent()) return extraction;
        if (stage1CompletedAtEpochMillis <= 0) return Optional.of("stage1CompletedAtEpochMillis が不正です");
        return Optional.empty();
    }

    private static Optional<String> validatePath(String field, String raw) {
        if (raw == null || raw.isBlank()) return Optional.of(field + " がありません");
        try {
            Path.of(raw);
            return Optional.empty();
        } catch (RuntimeException ex) {
            return Optional.of(field + " が不正です");
        }
    }

    private static String normalizePathText(String raw) {
        if (raw == null || raw.isBlank()) {
            return "";
        }
        try {
            return Path.of(raw.strip()).toAbsolutePath().normalize().toString();
        } catch (RuntimeException ex) {
            return raw.strip();
        }
    }
}
