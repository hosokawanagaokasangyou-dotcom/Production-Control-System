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
import java.util.regex.Pattern;

import jp.co.pm.ai.desktop.io.actuals.JuchuProcessingFeeRateLoader.FeeInfo;
import jp.co.pm.ai.desktop.ui.PlanInputProcessSequenceRowOrder;

/**
 * 加工賃トレンド用の日次円集計および依頼NO別一覧。
 *
 * <p>複数工程の依頼は受注「加工内容」末尾＝最終工程の m のみを使う。日次円は受注 AO（依頼の加工賃合計）を
 * 期間内最終工程 m 比率で按分する。AO 欠落時は AH（末尾単価）×最終工程 m。
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
    }

    /** 依頼NO別表の先頭合計行ラベル。 */
    public static final String TOTAL_REQUEST_LABEL = "合計";

    /**
     * 依頼NO別一覧の先頭に合計行を付ける。空なら空リスト。
     * 合計行の AO は按分円（実績円＋予定円）とし、各行 AO（受注額）の単純合算と混同しない。
     * 円/m は合算しない（欠落扱いで UI は —）。
     */
    public static List<RequestPoint> withLeadingTotalRow(List<RequestPoint> requests) {
        if (requests == null || requests.isEmpty()) {
            return List.of();
        }
        double actM = 0;
        double planM = 0;
        double actYen = 0;
        double planYen = 0;
        for (RequestPoint r : requests) {
            if (r == null || r.isTotalRow()) {
                continue;
            }
            actM += r.actualMeters();
            planM += r.planMeters();
            actYen += r.actualYen();
            planYen += r.planYen();
        }
        // 合計行 AO = 按分された実績円+予定円（受注月のみの未加工 AO は含めない）
        double allocatedAo = actYen + planYen;
        RequestPoint total =
                new RequestPoint(
                        TOTAL_REQUEST_LABEL, 0.0, allocatedAo, true, actM, planM, actYen, planYen);
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
     * 按分円合計（実績+予定）とは一致しない場合がある。
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

    /**
     * @param fees 依頼No → 受注 FeeInfo（AO・AH・加工内容）。null 可
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

        Map<String, Double> yenPerM = resolveYenPerMeter(actFiltered, planFiltered, feeMap);

        TreeMap<LocalDate, double[]> byDay = new TreeMap<>();
        for (LocalDate d = from; !d.isAfter(to); d = d.plusDays(1)) {
            byDay.put(d, new double[2]); // 0=actualYen, 1=planYen
        }
        // requestNo -> [actualM, planM, actualYen, planYen, rateMissingFlag(1/0), aoYen]
        TreeMap<String, double[]> byRequest = new TreeMap<>();

        int[] actCount = {0};
        int[] planCount = {0};
        int[] missing = {0};
        accumulate(actFiltered, yenPerM, feeMap, byDay, byRequest, 0, actCount, missing);
        accumulate(planFiltered, yenPerM, feeMap, byDay, byRequest, 1, planCount, missing);
        includeOrderMonthRequests(feeMap, byRequest, from, to);

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

        List<RequestPoint> requests = new ArrayList<>(byRequest.size());
        for (Map.Entry<String, double[]> e : byRequest.entrySet()) {
            double[] v = e.getValue();
            boolean missingRate = v[4] > 0.5;
            FeeInfo info = feeMap.get(e.getKey());
            double rateVal = info != null && info.hasAh() ? info.rateAhYenPerM() : 0.0;
            double aoVal = v[5];
            if (aoVal <= EPS && info != null && info.hasAo()) {
                aoVal = info.totalAoYen();
            }
            Double ypm = yenPerM.get(e.getKey());
            if (ypm != null && ypm > EPS) {
                rateVal = ypm;
            }
            requests.add(
                    new RequestPoint(
                            e.getKey(),
                            rateVal,
                            aoVal,
                            missingRate,
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

    /**
     * 期間に重なる受注年月（希望納期の年×月数）の依頼を依頼NO表へ含める。
     * 加工実績・予定が無い依頼も AO 合計が受注側と一致するようにする。
     */
    static void includeOrderMonthRequests(
            Map<String, FeeInfo> fees,
            TreeMap<String, double[]> byRequest,
            LocalDate from,
            LocalDate to) {
        if (fees == null || fees.isEmpty() || byRequest == null) {
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
            String req = e.getKey() == null || e.getKey().isBlank() ? "（依頼NOなし）" : e.getKey();
            double[] arr = byRequest.computeIfAbsent(req, k -> new double[6]);
            if (info.hasAo() && arr[5] <= EPS) {
                arr[5] = info.totalAoYen();
            }
        }
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

    /** 依頼NO → 円/m（AO÷期間最終工程m合計。AO無ければ AH）。 */
    static Map<String, Double> resolveYenPerMeter(
            List<QuantityLine> actualFinal,
            List<QuantityLine> planFinal,
            Map<String, FeeInfo> fees) {
        Map<String, Double> meters = new HashMap<>();
        addMeters(meters, actualFinal);
        addMeters(meters, planFinal);
        Map<String, Double> out = new LinkedHashMap<>();
        Set<String> keys = new HashSet<>();
        keys.addAll(meters.keySet());
        keys.addAll(fees.keySet());
        for (String req : keys) {
            FeeInfo info = fees.get(req);
            double m = meters.getOrDefault(req, 0.0);
            if (info != null && info.hasAo() && m > EPS) {
                out.put(req, info.totalAoYen() / m);
            } else if (info != null && info.hasAh()) {
                out.put(req, info.rateAhYenPerM());
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

    private static void accumulate(
            List<QuantityLine> lines,
            Map<String, Double> yenPerM,
            Map<String, FeeInfo> fees,
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
            Double ypm = yenPerM.get(line.requestNo());
            if (ypm == null) {
                ypm = yenPerM.get(req);
            }
            boolean missing = ypm == null;
            if (missing) {
                missingRate[0]++;
            } else {
                slotArr[slot] += line.meters() * ypm;
                counted[0]++;
            }
            double[] reqArr = byRequest.computeIfAbsent(req, k -> new double[6]);
            reqArr[slot] += line.meters();
            if (!missing) {
                reqArr[slot + 2] += line.meters() * ypm;
            } else {
                reqArr[4] = 1.0;
            }
            FeeInfo info = fees.get(line.requestNo());
            if (info != null && info.hasAo()) {
                reqArr[5] = info.totalAoYen();
            }
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
