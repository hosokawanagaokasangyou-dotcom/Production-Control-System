package jp.co.pm.ai.desktop.io.actuals;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.TreeMap;

/**
 * 加工賃トレンド用の日次円集計（AH × m）および依頼NO別一覧。
 *
 * <p>入力は日付・依頼No・m の行。単価マップに無い依頼は 0 円（日次も依頼別も）。
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

    /**
     * 見込累計は加工量トレンドと同型（当日までは実績、翌日以降は予定。先端で接続）。
     */
    public record DayPoint(
            LocalDate date,
            double actualYen,
            double planYen,
            double actualCumYen,
            double planCumYen,
            /** 見込累計（チャートで実績累計線と接続する破線）。 */
            double projectedCumYen) {}

    /** 期間内の依頼NO別集計行。 */
    public record RequestPoint(
            String requestNo,
            double rateYenPerM,
            boolean rateMissing,
            double actualMeters,
            double planMeters,
            double actualYen,
            double planYen) {
        public RequestPoint {
            requestNo = requestNo == null ? "" : requestNo.strip();
        }
    }

    public record Result(
            List<DayPoint> days,
            List<RequestPoint> requests,
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
        // requestNo -> [actualM, planM, actualYen, planYen, rateMissingFlag(1/0)]
        TreeMap<String, double[]> byRequest = new TreeMap<>();

        int[] actCount = {0};
        int[] planCount = {0};
        int[] missing = {0};
        accumulate(actualLines, rateMap, byDay, byRequest, 0, actCount, missing);
        accumulate(planLines, rateMap, byDay, byRequest, 1, planCount, missing);

        List<DayPoint> days = new ArrayList<>(byDay.size());
        double actCum = 0;
        double planCum = 0;
        double projCum = 0;
        double actTotal = 0;
        double planTotal = 0;
        for (Map.Entry<LocalDate, double[]> e : byDay.entrySet()) {
            LocalDate d = e.getKey();
            double a = e.getValue()[0];
            double p = e.getValue()[1];
            actTotal += a;
            planTotal += p;
            boolean usesPlan = d.isAfter(t);
            // 当日まで実績、翌日以降は予定（実績累計の先端と必ず接続）
            double projected = usesPlan ? p : a;
            if (!d.isAfter(t)) {
                actCum += a;
            }
            planCum += p;
            projCum += projected;
            days.add(new DayPoint(d, a, p, actCum, planCum, projCum));
        }

        List<RequestPoint> requests = new ArrayList<>(byRequest.size());
        for (Map.Entry<String, double[]> e : byRequest.entrySet()) {
            double[] v = e.getValue();
            boolean missingRate = v[4] > 0.5;
            Double rate = rateMap.get(e.getKey());
            double rateVal = rate != null ? rate : 0.0;
            requests.add(
                    new RequestPoint(
                            e.getKey(),
                            rateVal,
                            missingRate || rate == null,
                            v[0],
                            v[1],
                            v[2],
                            v[3]));
        }

        return new Result(
                List.copyOf(days),
                List.copyOf(requests),
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
            TreeMap<String, double[]> byRequest,
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
            String req = line.requestNo();
            if (req.isEmpty()) {
                req = "（依頼NOなし）";
            }
            Double rate = rates.get(line.requestNo());
            boolean missing = rate == null;
            if (missing) {
                missingRate[0]++;
            } else {
                slotArr[slot] += line.meters() * rate;
                counted[0]++;
            }
            double[] reqArr =
                    byRequest.computeIfAbsent(req, k -> new double[5]);
            reqArr[slot] += line.meters(); // 0=actualM, 1=planM
            if (!missing) {
                reqArr[slot + 2] += line.meters() * rate; // 2=actualYen, 3=planYen
            } else {
                reqArr[4] = 1.0;
            }
        }
    }
}
