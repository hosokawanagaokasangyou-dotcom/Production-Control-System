package jp.co.pm.ai.desktop.dispatch;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;

/**
 * 配台計画手動修正タブの日付列軸（横持ち・工程+機械×日）を決める。
 *
 * <p>段階3試行では計画暦日にタイムライン未割付の行を翌稼働日以降へスライドするが、
 * JSON の配台日 max だけを軸にするとスライド先の暦日列が無く UI 上で配台できない。
 * 未達（meta_miss）・納期・余白日を統合して min..max の連続範囲を返す。
 */
public final class DispatchInteractiveDateAxis {

    /**
     * 未達暦日の最大日から先へ、段階3スライド先を載せる暦日余白（稼働日カレンダー未参照の近似）。
     */
    public static final int SLIDE_CUSHION_CALENDAR_DAYS = 7;

    /** 日付列の開始を {@code today.minusDays(n)} まで広げる既定 n（UI 未設定時）。 */
    public static final int DEFAULT_DATE_AXIS_PAST_DAYS = 1;

    /**
     * @deprecated {@link #DEFAULT_DATE_AXIS_PAST_DAYS} を使用。
     */
    @Deprecated
    public static final int DISPATCH_WIDE_DATE_AXIS_PAST_DAYS = DEFAULT_DATE_AXIS_PAST_DAYS;

    private static final String COL_PROCESS_START = "加工開始日時";
    private static final String COL_PROCESS_COMPLETE = "加工完了日";
    private static final String COL_SPECIFIED_DUE = "指定納期";
    private static final String COL_ANSWER_DUE = "回答納期";

    private DispatchInteractiveDateAxis() {}

    /**
     * @param doc 結果_配台表 JSON
     * @param aladdinLookup shaped アラジン計画 lookup（空可）
     * @param trialShortfalls 試行直後の {@code dispatch_qty_shortfall}（空可）
     */
    public static List<LocalDate> computeInclusiveRange(
            ResultDispatchDocument doc,
            Map<String, Map<String, Map<String, Map<String, Double>>>> aladdinLookup,
            List<DispatchTrialShortages.DispatchQtyShortfallRow> trialShortfalls) {
        return computeInclusiveRange(
                doc, aladdinLookup, trialShortfalls, DEFAULT_DATE_AXIS_PAST_DAYS);
    }

    public static List<LocalDate> computeInclusiveRange(
            ResultDispatchDocument doc,
            Map<String, Map<String, Map<String, Map<String, Double>>>> aladdinLookup,
            List<DispatchTrialShortages.DispatchQtyShortfallRow> trialShortfalls,
            int pastDaysFromToday) {
        if (doc == null || doc.rows().isEmpty()) {
            return List.of();
        }
        TreeSet<LocalDate> ds = new TreeSet<>(ResultDispatchPivot.distinctDates(doc.rows()));
        collectAladdinPlanDatesInto(ds, doc.rows(), aladdinLookup);
        collectTaskDeadlineDatesInto(ds, doc.rows());
        collectShortfallDatesInto(ds, trialShortfalls);
        collectShortfallDatesInto(
                ds, DispatchTimelineMetaMissShortfalls.detectFromDocument(doc));
        List<DispatchTrialShortages.DispatchQtyShortfallRow> mergedShortfalls =
                mergeShortfallRows(
                        trialShortfalls,
                        DispatchTimelineMetaMissShortfalls.detectFromDocument(doc));
        collectSlideCushionInto(ds, doc, mergedShortfalls);
        if (ds.isEmpty()) {
            return defaultAxisWhenNoDataDates(pastDaysFromToday);
        }
        return extendAxisMinToPastDays(
                ResultDispatchPivot.dateRangeInclusive(new ArrayList<>(ds)),
                pastDaysFromToday);
    }

    /**
     * タスク×日付の日付列: 最小日を {@code today.minusDays(pastDays)} まで引き伸ばす（既にそれより過去ならそのまま）。
     */
    public static List<LocalDate> extendAxisMinToPastDays(List<LocalDate> axis, int pastDays) {
        if (axis == null || axis.isEmpty() || pastDays < 0) {
            return axis != null ? axis : List.of();
        }
        LocalDate floor = LocalDate.now().minusDays(pastDays);
        LocalDate min = axis.getFirst();
        if (!min.isAfter(floor)) {
            return axis;
        }
        return ResultDispatchPivot.dateRangeInclusive(List.of(floor, axis.getLast()));
    }

    /** JSON 等に配台日が無いときの既定軸（過去 {@link #DEFAULT_DATE_AXIS_PAST_DAYS} 日〜先の余白）。 */
    public static List<LocalDate> defaultAxisWhenNoDataDates() {
        return defaultAxisWhenNoDataDates(DEFAULT_DATE_AXIS_PAST_DAYS);
    }

    public static List<LocalDate> defaultAxisWhenNoDataDates(int pastDaysFromToday) {
        LocalDate today = LocalDate.now();
        LocalDate start = today.minusDays(Math.max(0, pastDaysFromToday));
        LocalDate end = today.plusDays(SLIDE_CUSHION_CALENDAR_DAYS);
        return ResultDispatchPivot.dateRangeInclusive(List.of(start, end));
    }

    /** デバッグ計測用: 各ソースが寄与した max 暦日。 */
    public static AxisExtent summarizeExtent(
            ResultDispatchDocument doc,
            Map<String, Map<String, Map<String, Map<String, Double>>>> aladdinLookup,
            List<DispatchTrialShortages.DispatchQtyShortfallRow> trialShortfalls) {
        if (doc == null || doc.rows().isEmpty()) {
            return new AxisExtent(null, null, null, null, null, null, 0);
        }
        LocalDate jsonMax =
                ResultDispatchPivot.distinctDates(doc.rows()).stream()
                        .max(LocalDate::compareTo)
                        .orElse(null);
        TreeSet<LocalDate> aladdinOnly = new TreeSet<>();
        collectAladdinPlanDatesInto(aladdinOnly, doc.rows(), aladdinLookup);
        LocalDate aladdinMax = aladdinOnly.isEmpty() ? null : aladdinOnly.last();

        TreeSet<LocalDate> deadlineOnly = new TreeSet<>();
        collectTaskDeadlineDatesInto(deadlineOnly, doc.rows());
        LocalDate deadlineMax = deadlineOnly.isEmpty() ? null : deadlineOnly.last();

        List<DispatchTrialShortages.DispatchQtyShortfallRow> merged =
                mergeShortfallRows(
                        trialShortfalls,
                        DispatchTimelineMetaMissShortfalls.detectFromDocument(doc));
        TreeSet<LocalDate> shortfallOnly = new TreeSet<>();
        collectShortfallDatesInto(shortfallOnly, merged);
        LocalDate shortfallMax = shortfallOnly.isEmpty() ? null : shortfallOnly.last();

        TreeSet<LocalDate> cushionOnly = new TreeSet<>(ResultDispatchPivot.distinctDates(doc.rows()));
        collectAladdinPlanDatesInto(cushionOnly, doc.rows(), aladdinLookup);
        collectTaskDeadlineDatesInto(cushionOnly, doc.rows());
        collectShortfallDatesInto(cushionOnly, merged);
        collectSlideCushionInto(cushionOnly, doc, merged);
        LocalDate axisMax = cushionOnly.isEmpty() ? null : cushionOnly.last();
        LocalDate axisMin = cushionOnly.isEmpty() ? null : cushionOnly.first();

        return new AxisExtent(
                jsonMax,
                aladdinMax,
                deadlineMax,
                shortfallMax,
                axisMin,
                axisMax,
                ResultDispatchPivot.dateRangeInclusive(new ArrayList<>(cushionOnly)).size());
    }

    public record AxisExtent(
            LocalDate jsonMax,
            LocalDate aladdinMax,
            LocalDate deadlineMax,
            LocalDate shortfallMax,
            LocalDate axisMin,
            LocalDate axisMax,
            int axisDayCount) {}

    private static void collectAladdinPlanDatesInto(
            Set<LocalDate> ds,
            List<Map<String, String>> rows,
            Map<String, Map<String, Map<String, Map<String, Double>>>> aladdinLookup) {
        if (ds == null || rows == null || rows.isEmpty() || aladdinLookup == null || aladdinLookup.isEmpty()) {
            return;
        }
        Set<String> seen = new HashSet<>();
        for (Map<String, String> row : rows) {
            String tid = nz(row.get("依頼NO"));
            String mk = nz(row.get(ResultDispatchSchema.COL_MACHINE));
            if (tid.isEmpty() || mk.isEmpty()) {
                continue;
            }
            String pair = tid + "\0" + mk;
            if (!seen.add(pair)) {
                continue;
            }
            for (LocalDate d : AladdinShapedPlanQtyLookup.distinctPlanDatesFor(aladdinLookup, mk, tid)) {
                ds.add(d);
            }
        }
    }

    private static void collectTaskDeadlineDatesInto(Set<LocalDate> ds, List<Map<String, String>> rows) {
        if (ds == null || rows == null) {
            return;
        }
        Set<String> seen = new HashSet<>();
        for (Map<String, String> row : rows) {
            String tid = nz(row.get("依頼NO"));
            String mk = nz(row.get(ResultDispatchSchema.COL_MACHINE));
            if (tid.isEmpty() || mk.isEmpty()) {
                continue;
            }
            if (!seen.add(tid + "\0" + mk)) {
                continue;
            }
            addParsedDateIfPresent(ds, row.get(COL_PROCESS_COMPLETE));
            addParsedDateIfPresent(ds, row.get(COL_SPECIFIED_DUE));
            addParsedDateIfPresent(ds, row.get(COL_ANSWER_DUE));
        }
    }

    private static void collectShortfallDatesInto(
            Set<LocalDate> ds, List<DispatchTrialShortages.DispatchQtyShortfallRow> rows) {
        if (ds == null || rows == null || rows.isEmpty()) {
            return;
        }
        for (DispatchTrialShortages.DispatchQtyShortfallRow r : rows) {
            if (r == null || r.dispatchDateIso() == null || r.dispatchDateIso().isBlank()) {
                continue;
            }
            try {
                ds.add(LocalDate.parse(r.dispatchDateIso().strip()));
            } catch (Exception ignored) {
                LocalDate d = ResultDispatchPivot.parseIsoDate(r.dispatchDateIso());
                if (d != null) {
                    ds.add(d);
                }
            }
        }
    }

    /**
     * 未達暦日（doc の meta_miss または shortfall 行）の最大日から先へ、段階3スライド先の余白を足す。
     */
    private static void collectSlideCushionInto(
            Set<LocalDate> ds,
            ResultDispatchDocument doc,
            List<DispatchTrialShortages.DispatchQtyShortfallRow> shortfalls) {
        if (ds == null) {
            return;
        }
        LocalDate maxAnchor = null;
        if (doc != null && !doc.rows().isEmpty()) {
            for (Map<String, String> row : doc.rows()) {
                if (row == null) {
                    continue;
                }
                double plan =
                        ResultDispatchNormalizer.parseDouble(
                                row.get(ResultDispatchSchema.COL_DISPATCH_QTY));
                if (plan <= 1e-6) {
                    continue;
                }
                if (!nz(row.get(COL_PROCESS_START)).isEmpty()) {
                    continue;
                }
                LocalDate dd =
                        ResultDispatchPivot.parseIsoDate(
                                row.get(ResultDispatchSchema.COL_DISPATCH_DATE));
                if (dd != null && (maxAnchor == null || dd.isAfter(maxAnchor))) {
                    maxAnchor = dd;
                }
            }
        }
        if (shortfalls != null) {
            for (DispatchTrialShortages.DispatchQtyShortfallRow r : shortfalls) {
                if (r == null || r.dispatchDateIso() == null || r.dispatchDateIso().isBlank()) {
                    continue;
                }
                LocalDate dd = null;
                try {
                    dd = LocalDate.parse(r.dispatchDateIso().strip());
                } catch (Exception ignored) {
                    dd = ResultDispatchPivot.parseIsoDate(r.dispatchDateIso());
                }
                if (dd != null && (maxAnchor == null || dd.isAfter(maxAnchor))) {
                    maxAnchor = dd;
                }
            }
        }
        if (maxAnchor == null) {
            return;
        }
        for (int i = 1; i <= SLIDE_CUSHION_CALENDAR_DAYS; i++) {
            ds.add(maxAnchor.plusDays(i));
        }
    }

    private static List<DispatchTrialShortages.DispatchQtyShortfallRow> mergeShortfallRows(
            List<DispatchTrialShortages.DispatchQtyShortfallRow> primary,
            List<DispatchTrialShortages.DispatchQtyShortfallRow> extra) {
        java.util.LinkedHashMap<String, DispatchTrialShortages.DispatchQtyShortfallRow> map =
                new java.util.LinkedHashMap<>();
        if (primary != null) {
            for (DispatchTrialShortages.DispatchQtyShortfallRow r : primary) {
                map.put(
                        DispatchTrialShortages.wideShortfallKey(
                                r.taskId(), r.machineName(), r.dispatchDateIso()),
                        r);
            }
        }
        if (extra != null) {
            for (DispatchTrialShortages.DispatchQtyShortfallRow r : extra) {
                map.putIfAbsent(
                        DispatchTrialShortages.wideShortfallKey(
                                r.taskId(), r.machineName(), r.dispatchDateIso()),
                        r);
            }
        }
        return List.copyOf(map.values());
    }

    private static void addParsedDateIfPresent(Set<LocalDate> ds, String raw) {
        LocalDate d = ResultDispatchPivot.parseIsoDate(raw);
        if (d != null) {
            ds.add(d);
        }
    }

    private static String nz(String s) {
        return s != null ? s.strip() : "";
    }
}
