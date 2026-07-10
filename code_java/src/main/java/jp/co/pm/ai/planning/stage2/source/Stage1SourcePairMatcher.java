package jp.co.pm.ai.planning.stage2.source;

import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;

/** 加工計画をアンカーに、同日の加工日報から時刻差最小のペアを組み立てる。 */
public final class Stage1SourcePairMatcher {

    public record MatchedPair(
            NetworkSourceExtractionCatalog.SourceEntry plan,
            NetworkSourceExtractionCatalog.SourceEntry dailyReport,
            long deltaMinutes,
            boolean sameDayDailyMissing,
            List<NetworkSourceExtractionCatalog.SourceEntry> sameDayDailyCandidates) {

        public MatchedPair {
            sameDayDailyCandidates =
                    sameDayDailyCandidates != null ? List.copyOf(sameDayDailyCandidates) : List.of();
        }

        public boolean largeDeltaWarning() {
            return NetworkSourceExtractionTimeSupport.isLargePairDelta(deltaMinutes);
        }
    }

    private Stage1SourcePairMatcher() {}

    public static List<MatchedPair> buildSelectableRows(Map<String, String> ui) {
        List<NetworkSourceExtractionCatalog.SourceEntry> plans =
                NetworkSourceExtractionCatalog.listProcessingPlans(ui);
        List<NetworkSourceExtractionCatalog.SourceEntry> dailyReports =
                NetworkSourceExtractionCatalog.listDailyReports(ui);
        List<MatchedPair> rows = new ArrayList<>();
        for (NetworkSourceExtractionCatalog.SourceEntry plan : plans) {
            rows.add(matchPlanToDaily(plan, dailyReports));
        }
        rows.sort(
                Comparator.comparing((MatchedPair r) -> r.plan().extractionTime())
                        .reversed());
        return List.copyOf(rows);
    }

    public static MatchedPair matchPlanToDaily(
            NetworkSourceExtractionCatalog.SourceEntry plan,
            List<NetworkSourceExtractionCatalog.SourceEntry> dailyReports) {
        Objects.requireNonNull(plan, "plan");
        LocalDate planDay = plan.extractionTime().toLocalDate();
        List<NetworkSourceExtractionCatalog.SourceEntry> sameDay = new ArrayList<>();
        for (NetworkSourceExtractionCatalog.SourceEntry dr : dailyReports) {
            if (dr.extractionTime().toLocalDate().equals(planDay)) {
                sameDay.add(dr);
            }
        }
        if (sameDay.isEmpty()) {
            return new MatchedPair(plan, null, Long.MAX_VALUE, true, List.of());
        }
        NetworkSourceExtractionCatalog.SourceEntry best =
                sameDay.stream()
                        .min(
                                Comparator.comparingLong(
                                        dr ->
                                                NetworkSourceExtractionTimeSupport.deltaMinutes(
                                                        plan.extractionTime(), dr.extractionTime())))
                        .orElse(null);
        long delta =
                best != null
                        ? NetworkSourceExtractionTimeSupport.deltaMinutes(
                                plan.extractionTime(), best.extractionTime())
                        : Long.MAX_VALUE;
        List<NetworkSourceExtractionCatalog.SourceEntry> candidates =
                sameDay.stream()
                        .sorted(
                                Comparator.comparing(
                                                NetworkSourceExtractionCatalog.SourceEntry::extractionTime)
                                        .reversed())
                        .toList();
        return new MatchedPair(plan, best, delta, false, candidates);
    }

    public static MatchedPair withDailyOverride(
            MatchedPair base, NetworkSourceExtractionCatalog.SourceEntry dailyOverride) {
        if (base == null || dailyOverride == null) {
            return base;
        }
        long delta =
                NetworkSourceExtractionTimeSupport.deltaMinutes(
                        base.plan().extractionTime(), dailyOverride.extractionTime());
        return new MatchedPair(
                base.plan(),
                dailyOverride,
                delta,
                false,
                base.sameDayDailyCandidates());
    }

    public static Optional<MatchedPair> findRowForPlanPath(
            Map<String, String> ui, Path planPath) {
        if (planPath == null) {
            return Optional.empty();
        }
        Path norm = planPath.toAbsolutePath().normalize();
        return buildSelectableRows(ui).stream()
                .filter(r -> r.plan().path().equals(norm))
                .findFirst();
    }
}
