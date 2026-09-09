package jp.co.pm.ai.desktop.io.actuals;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.TreeMap;

/**
 * 加工賃トレンド用の日次円集計（AH × m）。
 *
 * <p>入力は日付・依頼No・m の行。単価マップに無い依頼は 0 円。
 */
public final class ProcessingFeeTrendAggregator {

    static final int MAX_DAYS = 1000;
    private static final double EPS = 1e-9;

    private ProcessingFeeTrendAggregator() {}

    /** 実績または予定の数量行。 */
    public record QuantityLine(LocalDate date, String requestNo, double meters) {
        public QuantityLine {
            Objects.requireNonNull(date, "date");
            requestNo = requestNo == null ? "" : requestNo.strip();
        }
    }

    public record DayPoint(
            LocalDate date,
            double actualYen,
            double planYen,
            double actualCumYen,
            double planCumYen) {}

    public record Result(
            List<DayPoint> days,
            double actualTotalYen,
            double planTotalYen,
            LocalDate today,
            LocalDate from,
            LocalDate to,
            int actualLinesCounted,
            int planLinesCounted,
            int missingRateLines) {}

    public static Result aggregate(
            List<QuantityLine> actualLines,
            List<QuantityLine> planLines,
            Map<String, Double> rates,
            LocalDate from,
            LocalDate to,
            LocalDate today) {
        Objects.requireNonNull(from, "from");
        Objects.requireNonNull(to, "to");
        LocalDate t = today != null ? today : LocalDate.now();
        if (to.isBefore(from)) {
            LocalDate tmp = from;
            from = to;
            to = tmp;
        }
        if (from.plusDays(MAX_DAYS - 1).isBefore(to)) {
            to = from.plusDays(MAX_DAYS - 1);
        }
        Map<String, Double> rateMap = rates != null ? rates : Map.of();

        TreeMap<LocalDate, double[]> byDay = new TreeMap<>();
        for (LocalDate d = from; !d.isAfter(to); d = d.plusDays(1)) {
            byDay.put(d, new double[2]); // 0=actualYen, 1=planYen
        }

        int[] actCount = {0};
        int[] planCount = {0};
        int[] missing = {0};
        accumulate(actualLines, rateMap, byDay, 0, actCount, missing);
        accumulate(planLines, rateMap, byDay, 1, planCount, missing);

        List<DayPoint> days = new ArrayList<>(byDay.size());
        double actCum = 0;
        double planCum = 0;
        double actTotal = 0;
        double planTotal = 0;
        for (Map.Entry<LocalDate, double[]> e : byDay.entrySet()) {
            LocalDate d = e.getKey();
            double a = e.getValue()[0];
            double p = e.getValue()[1];
            actTotal += a;
            planTotal += p;
            if (!d.isAfter(t)) {
                actCum += a;
            }
            planCum += p;
            days.add(new DayPoint(d, a, p, actCum, planCum));
        }
        return new Result(
                List.copyOf(days),
                actTotal,
                planTotal,
                t,
                from,
                to,
                actCount[0],
                planCount[0],
                missing[0]);
    }

    private static void accumulate(
            List<QuantityLine> lines,
            Map<String, Double> rates,
            TreeMap<LocalDate, double[]> byDay,
            int slot,
            int[] counted,
            int[] missingRate) {
        if (lines == null) {
            return;
        }
        for (QuantityLine line : lines) {
            if (line == null || Math.abs(line.meters()) <= EPS) {
                continue;
            }
            double[] slotArr = byDay.get(line.date());
            if (slotArr == null) {
                continue;
            }
            Double rate = rates.get(line.requestNo());
            if (rate == null) {
                missingRate[0]++;
                continue;
            }
            slotArr[slot] += line.meters() * rate;
            counted[0]++;
        }
    }
}
