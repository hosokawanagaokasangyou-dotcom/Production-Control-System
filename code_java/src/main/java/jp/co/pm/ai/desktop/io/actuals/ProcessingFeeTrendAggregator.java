package jp.co.pm.ai.desktop.io.actuals;

import java.text.Normalizer;
import java.time.LocalDate;
import java.time.YearMonth;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.TreeMap;
import java.util.TreeSet;
import java.util.regex.Pattern;

import jp.co.pm.ai.desktop.io.actuals.JuchuProcessingFeeRateLoader.FeeInfo;
import jp.co.pm.ai.desktop.ui.PlanInputProcessSequenceRowOrder;

/**
 * 加工賃トレンド用の日次円集計および依頼NO別一覧。
 *
 * <p>AO（受注額）を正とする。円/m = AO ÷ 受注最終工程 m。実績円 = 円/m × 実績 m、
 * 未了（残予定）円 = 円/m × (受注 m − 実績 m) とし、両者の合計は当該依頼の AO に一致する。
 * 複数工程は実績・受注とも最終工程 m のみ。AO 欠落時は AH × m（未了は受注 m があれば同様）。
 */
public final class ProcessingFeeTrendAggregator {

    static final int MAX_DAYS = 1000;
    private static final double EPS = 1e-9;
    private static final Pattern WS = Pattern.compile("[\\s　]+");

    private ProcessingFeeTrendAggregator() {}

    /** 実績または予定の数量行。 */
    public record QuantityLine(LocalDate date, String requestNo, double meters, String processName) {
        public QuantityLine(LocalDate date, String requestNo, double meters) {
            this(date, requestNo, meters, "");
        }

        public QuantityLine {
            Objects.requireNonNull(date, "date");
            requestNo = requestNo == null ? "" : requestNo.strip();
            processName = processName == null ? "" : processName.strip();
        }
    }

    /**
     * 見込累計は加工量トレンドと同型（当日までは実績、翌日以降は予定。先端で接続）。
     * 翌日以降で同日に実績がある分は予定から差し引き（二重計上防止）。
     * 実績・予定・見込の累計はいずれも月初でリセットする。
     */
    public record DayPoint(
            LocalDate date,
            double actualYen,
            double planYen,
            double actualCumYen,
            double planCumYen,
            /** 見込累計（チャートで実績累計線と接続する破線）。月初リセット。 */
            double projectedCumYen) {}

    /** 期間内の依頼NO別集計行。 */
    public record RequestPoint(
            String requestNo,
            double rateYenPerM,
            double aoYen,
            boolean rateMissing,
            double actualMeters,
            double planMeters,
            double actualYen,
            double planYen) {
        public RequestPoint {
            requestNo = requestNo == null ? "" : requestNo.strip();
        }

        /** 互換: AO なしの旧コンストラクタ相当。 */
        public RequestPoint(
                String requestNo,
                double rateYenPerM,
                boolean rateMissing,
                double actualMeters,
                double planMeters,
                double actualYen,
                double planYen) {
            this(requestNo, rateYenPerM, 0.0, rateMissing, actualMeters, planMeters, actualYen, planYen);
        }

        public boolean isTotalRow() {
            return TOTAL_REQUEST_LABEL.equals(requestNo);
        }

        /** 未了 m（受注最終工程 m − 実績 m）。フィールド名 planMeters の別名。 */
        public double remainMeters() {
            return planMeters;
        }
    }

    /** 依頼NO別表の先頭合計行ラベル。 */
    public static final String TOTAL_REQUEST_LABEL = "合計";

    /**
     * 依頼NO別一覧の先頭に合計行を付ける。空なら空リスト。
     * AO 正本のとき、合計行の AO と実績円＋予定円（未了）は一致する（単価欠落行を除く）。
     * 円/m は合算しない（欠落扱いで UI は —）。
     */
    public static List<RequestPoint> withLeadingTotalRow(List<RequestPoint> requests) {
        if (requests == null || requests.isEmpty()) {
            return List.of();
        }
        double ao = 0;
        double actM = 0;
        double planM = 0;
        double actYen = 0;
        double planYen = 0;
        for (RequestPoint r : requests) {
            if (r == null || r.isTotalRow()) {
                continue;
            }
            ao += r.aoYen();
            actM += r.actualMeters();
            planM += r.planMeters();
            actYen += r.actualYen();
            planYen += r.planYen();
        }
        RequestPoint total =
                new RequestPoint(TOTAL_REQUEST_LABEL, 0.0, ao, true, actM, planM, actYen, planYen);
        List<RequestPoint> out = new ArrayList<>(requests.size() + 1);
        out.add(total);
        for (RequestPoint r : requests) {
            if (r != null && !r.isTotalRow()) {
                out.add(r);
            }
        }
        return List.copyOf(out);
    }

    /**
     * 一覧上の受注 AO 合計（受注年月で含めた未加工依頼分も含む）。
     * {@link #withLeadingTotalRow} の合計行 AO と同値。
     */
    public static double sumOrderAoYen(List<RequestPoint> requests) {
        if (requests == null || requests.isEmpty()) {
            return 0.0;
        }
        double ao = 0;
        for (RequestPoint r : requests) {
            if (r == null || r.isTotalRow()) {
                continue;
            }
            ao += r.aoYen();
        }
        return ao;
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

    /** 依頼ごとの AO 正本配分結果。 */
    record RequestAlloc(
            double yenPerM,
            double orderMeters,
            double actualMeters,
            double remainMeters,
            double actualYen,
            double planYen,
            double aoYen,
            boolean rateMissing) {}

    /**
     * @param fees 依頼No → 受注 FeeInfo（AO・AH・加工内容・受注最終工程 m）。null 可
     */
    public static Result aggregate(
            List<QuantityLine> actualLines,
            List<QuantityLine> planLines,
            Map<String, FeeInfo> fees,
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
        Map<String, FeeInfo> feeMap = fees != null ? fees : Map.of();

        List<QuantityLine> actFiltered = filterToFinalProcess(actualLines, feeMap);
        List<QuantityLine> planFiltered = filterToFinalProcess(planLines, feeMap);

        Map<String, Double> actualMetersByReq = sumMetersByRequest(actFiltered);
        Set<String> reqKeys = new HashSet<>(actualMetersByReq.keySet());
        addOrderMonthRequestKeys(feeMap, reqKeys, from, to);
        for (QuantityLine line : planFiltered) {
            if (line != null && !line.requestNo().isEmpty()) {
                reqKeys.add(line.requestNo());
            }
        }

        Map<String, RequestAlloc> allocs = new LinkedHashMap<>();
        int missing = 0;
        for (String req : new TreeSet<>(reqKeys)) {
            FeeInfo info = feeMap.get(req);
            double actM = actualMetersByReq.getOrDefault(req, 0.0);
            RequestAlloc alloc = allocateRequest(info, actM, planFiltered, req);
            allocs.put(req, alloc);
            if (alloc.rateMissing()) {
                missing++;
            }
        }

        TreeMap<LocalDate, double[]> byDay = new TreeMap<>();
        for (LocalDate d = from; !d.isAfter(to); d = d.plusDays(1)) {
            byDay.put(d, new double[2]);
        }
        int actCount =
                accumulateActualDays(actFiltered, allocs, feeMap, byDay);
        int planCount =
                distributeRemainToPlanDays(planFiltered, allocs, byDay, from, to, t);

        List<DayPoint> days = new ArrayList<>(byDay.size());
        double actCum = 0;
        double planCum = 0;
        double projCum = 0;
        double actTotal = 0;
        double planTotal = 0;
        YearMonth cumMonth = null;
        for (Map.Entry<LocalDate, double[]> e : byDay.entrySet()) {
            LocalDate d = e.getKey();
            YearMonth ym = YearMonth.from(d);
            if (cumMonth == null || !ym.equals(cumMonth)) {
                actCum = 0;
                planCum = 0;
                projCum = 0;
                cumMonth = ym;
            }
            double a = e.getValue()[0];
            double p = e.getValue()[1];
            boolean usesPlan = d.isAfter(t);
            double planForMetrics = usesPlan ? Math.max(0.0, p - a) : p;
            actTotal += a;
            planTotal += planForMetrics;
            double projected = usesPlan ? planForMetrics : a;
            if (!d.isAfter(t)) {
                actCum += a;
            }
            planCum += planForMetrics;
            projCum += projected;
            days.add(new DayPoint(d, a, planForMetrics, actCum, planCum, projCum));
        }

        List<RequestPoint> requests = new ArrayList<>(allocs.size());
        for (Map.Entry<String, RequestAlloc> e : allocs.entrySet()) {
            RequestAlloc al = e.getValue();
            requests.add(
                    new RequestPoint(
                            e.getKey(),
                            al.yenPerM(),
                            al.aoYen(),
                            al.rateMissing(),
                            al.actualMeters(),
                            al.remainMeters(),
                            al.actualYen(),
                            al.planYen()));
        }

        return new Result(
                List.copyOf(days),
                List.copyOf(requests),
                actTotal,
                planTotal,
                t,
                from,
                to,
                actCount,
                planCount,
                missing);
    }

    static RequestAlloc allocateRequest(
            FeeInfo info, double actualMeters, List<QuantityLine> planFiltered, String req) {
        double actM = Math.max(0.0, actualMeters);
        double ao = info != null && info.hasAo() ? info.totalAoYen() : 0.0;
        if (info != null && info.hasAo() && info.hasOrderFinalMeters()) {
            double orderM = info.orderFinalMeters();
            double ypm = ao / orderM;
            double actUsed = Math.min(actM, orderM);
            double remain = Math.max(0.0, orderM - actUsed);
            return new RequestAlloc(
                    ypm, orderM, actUsed, remain, actUsed * ypm, remain * ypm, ao, false);
        }
        if (info != null && info.hasAo() && !info.hasOrderFinalMeters()) {
            // 受注 m 不明: 期間実績が無ければ全額を未了、あれば実績に全額（一致優先）
            if (actM <= EPS) {
                return new RequestAlloc(0.0, 0.0, 0.0, 0.0, 0.0, ao, ao, false);
            }
            double ypm = ao / actM;
            return new RequestAlloc(ypm, actM, actM, 0.0, ao, 0.0, ao, false);
        }
        if (info != null && info.hasAh()) {
            double ypm = info.rateAhYenPerM();
            double planSched = sumMetersForRequest(planFiltered, req);
            double orderM =
                    info.hasOrderFinalMeters() ? info.orderFinalMeters() : actM + planSched;
            if (orderM <= EPS) {
                orderM = actM;
            }
            double actUsed = orderM > EPS ? Math.min(actM, orderM) : actM;
            double remain = Math.max(0.0, orderM - actUsed);
            return new RequestAlloc(
                    ypm,
                    orderM,
                    actUsed,
                    remain,
                    actUsed * ypm,
                    remain * ypm,
                    ao,
                    false);
        }
        return new RequestAlloc(0.0, 0.0, actM, 0.0, 0.0, 0.0, ao, actM > EPS || ao > EPS);
    }

    private static double sumMetersForRequest(List<QuantityLine> lines, String req) {
        if (lines == null || req == null) {
            return 0.0;
        }
        double s = 0;
        for (QuantityLine line : lines) {
            if (line != null && req.equals(line.requestNo())) {
                s += line.meters();
            }
        }
        return s;
    }

    private static Map<String, Double> sumMetersByRequest(List<QuantityLine> lines) {
        Map<String, Double> meters = new HashMap<>();
        addMeters(meters, lines);
        return meters;
    }

    private static void addOrderMonthRequestKeys(
            Map<String, FeeInfo> fees, Set<String> reqKeys, LocalDate from, LocalDate to) {
        if (fees == null || fees.isEmpty()) {
            return;
        }
        Set<YearMonth> months = new HashSet<>();
        for (LocalDate d = from; !d.isAfter(to); d = d.plusDays(1)) {
            months.add(YearMonth.from(d));
        }
        for (Map.Entry<String, FeeInfo> e : fees.entrySet()) {
            FeeInfo info = e.getValue();
            if (info == null || !info.hasOrderYearMonth()) {
                continue;
            }
            YearMonth ym = YearMonth.of(info.orderYear(), info.orderMonth());
            if (!months.contains(ym)) {
                continue;
            }
            if (e.getKey() != null && !e.getKey().isBlank()) {
                reqKeys.add(e.getKey());
            }
        }
    }

    /** 実績日次円。受注 m を超えないよう依頼ごとにクリップ。 */
    private static int accumulateActualDays(
            List<QuantityLine> actFiltered,
            Map<String, RequestAlloc> allocs,
            Map<String, FeeInfo> feeMap,
            TreeMap<LocalDate, double[]> byDay) {
        if (actFiltered == null) {
            return 0;
        }
        Map<String, Double> used = new HashMap<>();
        int counted = 0;
        for (QuantityLine line : actFiltered) {
            if (line == null || Math.abs(line.meters()) <= EPS) {
                continue;
            }
            double[] slotArr = byDay.get(line.date());
            if (slotArr == null) {
                continue;
            }
            String req = line.requestNo().isEmpty() ? "（依頼NOなし）" : line.requestNo();
            RequestAlloc al = allocs.get(req);
            if (al == null || al.rateMissing() || al.yenPerM() <= EPS) {
                continue;
            }
            double already = used.getOrDefault(req, 0.0);
            double room =
                    al.orderMeters() > EPS
                            ? Math.max(0.0, al.orderMeters() - already)
                            : line.meters();
            double take = Math.min(line.meters(), room);
            if (take <= EPS) {
                continue;
            }
            slotArr[0] += take * al.yenPerM();
            used.put(req, already + take);
            counted++;
        }
        return counted;
    }

    /**
     * 未了円を予定日に配分。翌日以降の予定 m 比率。予定行が無ければ期間末（または翌日）へ一括。
     */
    private static int distributeRemainToPlanDays(
            List<QuantityLine> planFiltered,
            Map<String, RequestAlloc> allocs,
            TreeMap<LocalDate, double[]> byDay,
            LocalDate from,
            LocalDate to,
            LocalDate today) {
        int counted = 0;
        for (Map.Entry<String, RequestAlloc> e : allocs.entrySet()) {
            RequestAlloc al = e.getValue();
            if (al.planYen() <= EPS || al.rateMissing()) {
                continue;
            }
            String req = e.getKey();
            Map<LocalDate, Double> weights = new HashMap<>();
            double wSum = 0;
            if (planFiltered != null) {
                for (QuantityLine line : planFiltered) {
                    if (line == null || !req.equals(line.requestNo())) {
                        continue;
                    }
                    if (!line.date().isAfter(today)) {
                        continue;
                    }
                    if (!byDay.containsKey(line.date())) {
                        continue;
                    }
                    weights.merge(line.date(), line.meters(), Double::sum);
                    wSum += line.meters();
                }
            }
            if (wSum > EPS) {
                for (Map.Entry<LocalDate, Double> w : weights.entrySet()) {
                    double yen = al.planYen() * (w.getValue() / wSum);
                    byDay.get(w.getKey())[1] += yen;
                }
                counted++;
            } else {
                LocalDate anchor = today.plusDays(1);
                if (anchor.isBefore(from)) {
                    anchor = from;
                }
                if (anchor.isAfter(to)) {
                    anchor = to;
                }
                byDay.get(anchor)[1] += al.planYen();
                counted++;
            }
        }
        return counted;
    }

    /**
     * 最終工程以外の行を落とす。加工内容が空で工程が複数ある依頼は全行落とす（誤合算防止）。
     */
    static List<QuantityLine> filterToFinalProcess(
            List<QuantityLine> lines, Map<String, FeeInfo> fees) {
        if (lines == null || lines.isEmpty()) {
            return List.of();
        }
        Map<String, Set<String>> processesByReq = new HashMap<>();
        for (QuantityLine line : lines) {
            if (line == null) {
                continue;
            }
            String req = line.requestNo().isEmpty() ? "（依頼NOなし）" : line.requestNo();
            String proc = normalizeProcessName(line.processName());
            if (!proc.isEmpty()) {
                processesByReq.computeIfAbsent(req, k -> new HashSet<>()).add(proc);
            }
        }
        List<QuantityLine> out = new ArrayList<>();
        for (QuantityLine line : lines) {
            if (line == null || Math.abs(line.meters()) <= EPS) {
                continue;
            }
            String reqKey = line.requestNo();
            FeeInfo info = fees.get(reqKey);
            String content = info != null ? info.processContent() : "";
            List<String> tokens = PlanInputProcessSequenceRowOrder.parseProcessContentTokens(content);
            String proc = normalizeProcessName(line.processName());
            if (tokens.size() >= 2) {
                String finalTok =
                        normalizeProcessName(tokens.get(tokens.size() - 1));
                if (finalTok.isEmpty() || !finalTok.equals(proc)) {
                    continue;
                }
            } else if (tokens.isEmpty()) {
                String lookup = reqKey.isEmpty() ? "（依頼NOなし）" : reqKey;
                Set<String> procs = processesByReq.getOrDefault(lookup, Set.of());
                if (procs.size() >= 2) {
                    // 加工内容無しで複数工程 → 採用しない
                    continue;
                }
            }
            // 単工程トークン or 工程1種類 or 工程名空: 採用
            out.add(line);
        }
        return out;
    }

    private static void addMeters(Map<String, Double> meters, List<QuantityLine> lines) {
        if (lines == null) {
            return;
        }
        for (QuantityLine line : lines) {
            if (line == null || Math.abs(line.meters()) <= EPS) {
                continue;
            }
            String req = line.requestNo().isEmpty() ? "（依頼NOなし）" : line.requestNo();
            meters.merge(req, line.meters(), Double::sum);
        }
    }

    static String normalizeProcessName(String raw) {
        if (raw == null) {
            return "";
        }
        String t = Normalizer.normalize(raw.strip(), Normalizer.Form.NFKC);
        return WS.matcher(t).replaceAll("");
    }
}
