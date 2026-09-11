package jp.co.pm.ai.desktop.io.actuals;

import java.text.Normalizer;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.YearMonth;
import java.util.ArrayList;
import java.util.Collections;
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
 * 依頼NO別の実績 m は表示期間外の日報出来高も含む（日次棒のみ表示期間内）。
 * 複数工程は実績・受注とも最終工程 m のみ。日報の最終工程は加工日付＋終了時間の最遅行の工程名。
 * AO 欠落時は AH × m（未了は受注 m があれば同様）。
 */
public final class ProcessingFeeTrendAggregator {

    static final int MAX_DAYS = 1000;
    private static final double EPS = 1e-9;
    private static final Pattern WS = Pattern.compile("[\\s　]+");

    private ProcessingFeeTrendAggregator() {}

    /**
     * 実績または予定の数量行。
     *
     * @param finishedAt 日報の加工日付＋終了時間。無い・未解析は {@code null}
     */
    public record QuantityLine(
            LocalDate date,
            String requestNo,
            double meters,
            String processName,
            LocalDateTime finishedAt) {
        public QuantityLine(LocalDate date, String requestNo, double meters) {
            this(date, requestNo, meters, "", null);
        }

        public QuantityLine(LocalDate date, String requestNo, double meters, String processName) {
            this(date, requestNo, meters, processName, null);
        }

        public QuantityLine {
            Objects.requireNonNull(date, "date");
            requestNo =
                    requestNo == null || requestNo.isBlank()
                            ? ""
                            : ProcessingTrendAggregator.normKey(requestNo);
            processName = processName == null ? "" : processName.strip();
        }
    }

    /**
     * 見込累計は加工量トレンドと同型（当日までは実績、翌日以降は未了。先端で接続）。
     * 翌日以降で同日に実績がある分は未了から差し引き（二重計上防止）。
     * 実績・未了・見込の累計はいずれも月初でリセットする。
     * 表示期間がすべて過去（本日以前）のときは日次の未了棒を出さない（KPI・依頼表の未了のみ）。
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
        Map<String, FeeInfo> feeMap = normalizeFeeKeys(fees);

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
        // 日次棒は表示期間内のみ。依頼NO別の実績 m は期間外の出来高も含む（AO 正本）
        List<QuantityLine> actInPeriod = new ArrayList<>();
        for (QuantityLine line : actFiltered) {
            if (line != null && !line.date().isBefore(from) && !line.date().isAfter(to)) {
                actInPeriod.add(line);
            }
        }
        int actCount = accumulateActualDays(actInPeriod, allocs, feeMap, byDay);
        boolean periodFullyPast = to.isBefore(t);
        // 過去期間: 日次チャートに未了棒を載せない（KPI・依頼表の未了は alloc 側）
        int planCount =
                periodFullyPast
                        ? 0
                        : distributeRemainToPlanDays(planFiltered, allocs, byDay, from, to, t);

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
            double p = periodFullyPast ? 0.0 : e.getValue()[1];
            boolean usesPlan = !periodFullyPast && d.isAfter(t);
            double planForMetrics = usesPlan ? Math.max(0.0, p - a) : p;
            actTotal += a;
            planTotal += planForMetrics;
            double projected = usesPlan ? planForMetrics : a;
            if (!d.isAfter(t) || periodFullyPast) {
                actCum += a;
            }
            planCum += planForMetrics;
            projCum += projected;
            days.add(new DayPoint(d, a, planForMetrics, actCum, planCum, projCum));
        }

        List<RequestPoint> requests = new ArrayList<>(allocs.size());
        double remainYenFromAlloc = 0;
        for (Map.Entry<String, RequestAlloc> e : allocs.entrySet()) {
            RequestAlloc al = e.getValue();
            remainYenFromAlloc += al.planYen();
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
        // 過去期間は日次未了棒を出さないため、KPI 未了は依頼配分の合計を使う
        if (periodFullyPast) {
            planTotal = remainYenFromAlloc;
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
     * 未了円を予定日に配分（進行中・未来の表示期間向け）。
     * 翌日以降の予定 m 比率。予定行が無ければ「今日の翌日」（期間外なら期間末）へ一括。
     * 過去期間のみの表示では呼ばない（日次未了棒は出さない）。
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

    private static Map<String, FeeInfo> normalizeFeeKeys(Map<String, FeeInfo> fees) {
        if (fees == null || fees.isEmpty()) {
            return Map.of();
        }
        Map<String, FeeInfo> out = new LinkedHashMap<>();
        for (Map.Entry<String, FeeInfo> e : fees.entrySet()) {
            if (e.getKey() == null || e.getKey().isBlank() || e.getValue() == null) {
                continue;
            }
            String k = ProcessingTrendAggregator.normKey(e.getKey());
            if (k.isEmpty()) {
                continue;
            }
            out.put(k, e.getValue());
        }
        return Collections.unmodifiableMap(out);
    }

    /**
     * 最終工程以外の行を落とす。
     *
     * <p>日報行に終了日時がある依頼は、加工日付＋終了時間の最遅行の工程名のみ採用する。
     * 終了日時が無い場合は受注「加工内容」末尾。加工内容が空で工程が複数ある依頼は全行落とす（誤合算防止）。
     */
    static List<QuantityLine> filterToFinalProcess(
            List<QuantityLine> lines, Map<String, FeeInfo> fees) {
        if (lines == null || lines.isEmpty()) {
            return List.of();
        }
        Map<String, List<QuantityLine>> byReq = new LinkedHashMap<>();
        Map<String, Set<String>> processesByReq = new HashMap<>();
        for (QuantityLine line : lines) {
            if (line == null || Math.abs(line.meters()) <= EPS) {
                continue;
            }
            String lookup = line.requestNo().isEmpty() ? "（依頼NOなし）" : line.requestNo();
            byReq.computeIfAbsent(lookup, k -> new ArrayList<>()).add(line);
            String proc = normalizeProcessName(line.processName());
            if (!proc.isEmpty()) {
                processesByReq.computeIfAbsent(lookup, k -> new HashSet<>()).add(proc);
            }
        }
        List<QuantityLine> out = new ArrayList<>();
        for (Map.Entry<String, List<QuantityLine>> e : byReq.entrySet()) {
            String lookup = e.getKey();
            List<QuantityLine> group = e.getValue();
            String byFinished = resolveFinalProcessByFinishedAt(group);
            if (byFinished != null) {
                for (QuantityLine line : group) {
                    if (byFinished.equals(normalizeProcessName(line.processName()))) {
                        out.add(line);
                    }
                }
                continue;
            }
            String reqKey = "（依頼NOなし）".equals(lookup) ? "" : lookup;
            FeeInfo info = fees.get(reqKey);
            String content = info != null ? info.processContent() : "";
            List<String> tokens = PlanInputProcessSequenceRowOrder.parseProcessContentTokens(content);
            for (QuantityLine line : group) {
                String proc = normalizeProcessName(line.processName());
                if (tokens.size() >= 2) {
                    String finalTok = normalizeProcessName(tokens.get(tokens.size() - 1));
                    if (finalTok.isEmpty() || !finalTok.equals(proc)) {
                        continue;
                    }
                } else if (tokens.isEmpty()) {
                    Set<String> procs = processesByReq.getOrDefault(lookup, Set.of());
                    if (procs.size() >= 2) {
                        // 加工内容無しで複数工程 → 採用しない（後段で受注 m 一致フォールバック）
                        continue;
                    }
                }
                out.add(line);
            }
        }
        return withOrderMetersProcessFallback(lines, out, fees);
    }

    /**
     * 依頼内で加工日付＋終了時間が最も遅い行の工程名。終了日時が無い・同刻で工程が複数なら null。
     */
    static String resolveFinalProcessByFinishedAt(List<QuantityLine> group) {
        if (group == null || group.isEmpty()) {
            return null;
        }
        LocalDateTime max = null;
        String bestProc = null;
        boolean ambiguous = false;
        for (QuantityLine line : group) {
            if (line == null || line.finishedAt() == null) {
                continue;
            }
            String proc = normalizeProcessName(line.processName());
            if (proc.isEmpty()) {
                continue;
            }
            LocalDateTime fa = line.finishedAt();
            if (max == null || fa.isAfter(max)) {
                max = fa;
                bestProc = proc;
                ambiguous = false;
            } else if (fa.equals(max) && !proc.equals(bestProc)) {
                ambiguous = true;
            }
        }
        if (max == null || ambiguous) {
            return null;
        }
        return bestProc;
    }

    /**
     * 最終工程名で 0 m になったとき、受注最終工程 m と一致する工程の出来高があればそれを採用する。
     * （例: 加工内容末尾が「増刷」で出来高 0、SEC が 4000=受注 m）
     */
    private static List<QuantityLine> withOrderMetersProcessFallback(
            List<QuantityLine> all, List<QuantityLine> filtered, Map<String, FeeInfo> fees) {
        Map<String, Double> filteredMeters = sumMetersByRequest(filtered);
        Map<String, List<QuantityLine>> byReq = new HashMap<>();
        for (QuantityLine line : all) {
            if (line == null || Math.abs(line.meters()) <= EPS || line.requestNo().isEmpty()) {
                continue;
            }
            byReq.computeIfAbsent(line.requestNo(), k -> new ArrayList<>()).add(line);
        }
        List<QuantityLine> out = new ArrayList<>(filtered);
        for (Map.Entry<String, List<QuantityLine>> e : byReq.entrySet()) {
            String req = e.getKey();
            if (filteredMeters.getOrDefault(req, 0.0) > EPS) {
                continue;
            }
            FeeInfo info = fees.get(req);
            if (info == null || !info.hasOrderFinalMeters()) {
                continue;
            }
            double orderM = info.orderFinalMeters();
            Map<String, Double> byProc = new HashMap<>();
            for (QuantityLine line : e.getValue()) {
                String proc = normalizeProcessName(line.processName());
                if (proc.isEmpty()) {
                    continue;
                }
                byProc.merge(proc, line.meters(), Double::sum);
            }
            String matchProc = null;
            for (Map.Entry<String, Double> p : byProc.entrySet()) {
                if (Math.abs(p.getValue() - orderM) <= Math.max(0.5, orderM * 1e-6)) {
                    if (matchProc != null) {
                        matchProc = null; // 複数工程が一致 → 曖昧なので使わない
                        break;
                    }
                    matchProc = p.getKey();
                }
            }
            if (matchProc == null) {
                continue;
            }
            for (QuantityLine line : e.getValue()) {
                if (matchProc.equals(normalizeProcessName(line.processName()))) {
                    out.add(line);
                }
            }
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
