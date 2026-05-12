package jp.co.pm.ai.desktop.dispatch;

import java.time.LocalDate;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;

/**
 * Builds pivot views from canonical long-format rows.
 */
public final class ResultDispatchPivot {

    private ResultDispatchPivot() {}

    /**
     * 「タスク×日付」ワイド表で同一タスクとみなす列（配台試行順・メンバー名・加工開始終了など日別メタは含めない）。
     */
    public static final List<String> DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS =
            List.of(
                    ResultDispatchSchema.COL_PROCESS,
                    ResultDispatchSchema.COL_MACHINE,
                    "加工内容",
                    "依頼NO",
                    "換算数量",
                    "計画合計");

    /** Distinct task rows (static columns only), insertion order. */
    public static List<Map<String, String>> distinctTaskProfiles(List<String> columns, List<Map<String, String>> rows) {
        String dc = ResultDispatchSchema.COL_DISPATCH_DATE;
        String qc = ResultDispatchSchema.COL_DISPATCH_QTY;
        Set<String> seen = new LinkedHashSet<>();
        List<Map<String, String>> out = new ArrayList<>();
        for (Map<String, String> row : rows) {
            String gk = ResultDispatchNormalizer.groupKey(columns, row, dc, qc);
            if (seen.add(gk)) {
                Map<String, String> profile = new LinkedHashMap<>();
                for (String col : columns) {
                    if (col.equals(dc) || col.equals(qc)) {
                        continue;
                    }
                    profile.put(col, nz(row.get(col)));
                }
                out.add(profile);
            }
        }
        return out;
    }

    /**
     * 「タスク×日付」用: {@link #DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS} が一致する行は同一プロファイルにまとめる。
     */
    public static List<Map<String, String>> distinctWideTaskProfiles(
            List<String> columns, List<Map<String, String>> rows, List<String> mergeIdentityHeaders) {
        String dc = ResultDispatchSchema.COL_DISPATCH_DATE;
        String qc = ResultDispatchSchema.COL_DISPATCH_QTY;
        Set<String> seen = new LinkedHashSet<>();
        List<Map<String, String>> out = new ArrayList<>();
        for (Map<String, String> row : rows) {
            String gk = wideMergeIdentityKey(row, mergeIdentityHeaders);
            if (seen.add(gk)) {
                LinkedHashMap<String, String> profile = new LinkedHashMap<>();
                for (String col : columns) {
                    if (col.equals(dc) || col.equals(qc)) {
                        continue;
                    }
                    profile.put(col, nz(row.get(col)));
                }
                out.add(profile);
            }
        }
        return out;
    }

    /**
     * 同一タスク（ワイド同一性）かつ同一配台日の行を 1 行にまとめ、数量を合算する。読込直後・再描画前に呼ぶ。
     */
    public static void mergeDispatchRowsByWideIdentity(
            List<String> columns, List<Map<String, String>> rows, List<String> mergeIdentityHeaders) {
        if (rows.isEmpty()) {
            return;
        }
        String dc = ResultDispatchSchema.COL_DISPATCH_DATE;
        String qc = ResultDispatchSchema.COL_DISPATCH_QTY;
        Map<String, Map<String, String>> buckets = new LinkedHashMap<>();
        for (Map<String, String> row : new ArrayList<>(rows)) {
            String mid = wideMergeIdentityKey(row, mergeIdentityHeaders);
            String dk = nz(row.get(dc));
            String key = mid + "\u0000" + dk;
            Map<String, String> existing = buckets.get(key);
            if (existing == null) {
                buckets.put(key, new LinkedHashMap<>(row));
            } else {
                double q1 = ResultDispatchNormalizer.parseDouble(existing.get(qc));
                double q2 = ResultDispatchNormalizer.parseDouble(row.get(qc));
                existing.put(qc, ResultDispatchNormalizer.formatQty(q1 + q2));
            }
        }
        rows.clear();
        rows.addAll(buckets.values());
    }

    static String wideMergeIdentityKey(Map<String, String> row, List<String> mergeIdentityHeaders) {
        StringBuilder sb = new StringBuilder();
        for (String h : mergeIdentityHeaders) {
            sb.append('\u0001');
            sb.append(nz(row.get(h)));
        }
        return sb.toString();
    }

    /** {@code mergeIdentityHeaders} に列挙した項目がすべて一致するか（ワイド「同一タスク」判定）。 */
    public static boolean matchesWideMergeIdentity(
            Map<String, String> profile, Map<String, String> row, List<String> mergeIdentityHeaders) {
        for (String h : mergeIdentityHeaders) {
            if (!nz(profile.get(h)).equals(nz(row.get(h)))) {
                return false;
            }
        }
        return true;
    }

    public static List<LocalDate> distinctDates(List<Map<String, String>> rows) {
        String dc = ResultDispatchSchema.COL_DISPATCH_DATE;
        Set<LocalDate> ds = new TreeSet<>();
        for (Map<String, String> row : rows) {
            LocalDate d = parseIsoDate(row.get(dc));
            if (d != null) {
                ds.add(d);
            }
        }
        return new ArrayList<>(ds);
    }

    /** Expand date range: min..max from rows, optionally extend empty grid (fill gap dates). */
    public static List<LocalDate> dateRangeInclusive(List<LocalDate> sortedDistinct) {
        if (sortedDistinct.isEmpty()) {
            return List.of();
        }
        LocalDate min = sortedDistinct.getFirst();
        LocalDate max = sortedDistinct.getLast();
        List<LocalDate> out = new ArrayList<>();
        for (LocalDate d = min; !d.isAfter(max); d = d.plusDays(1)) {
            out.add(d);
        }
        return out;
    }

    public static double sumQuantityForProfileAndDate(
            List<String> columns,
            List<Map<String, String>> rows,
            Map<String, String> taskProfile,
            LocalDate date) {
        String dc = ResultDispatchSchema.COL_DISPATCH_DATE;
        String qc = ResultDispatchSchema.COL_DISPATCH_QTY;
        double sum = 0;
        String ds = date.toString();
        for (Map<String, String> row : rows) {
            if (!profileMatches(columns, taskProfile, row, dc, qc)) {
                continue;
            }
            String rd = nz(row.get(dc));
            if (!ds.equals(rd) && !ds.equals(normalizeDateCell(rd))) {
                continue;
            }
            sum += ResultDispatchNormalizer.parseDouble(row.get(qc));
        }
        return sum;
    }

    /** {@link #sumQuantityForProfileAndDate} のワイド版: メタ列の不一致で同一タスクが分断されていても合算する。 */
    public static double sumQuantityForProfileAndDateForWideMerge(
            List<Map<String, String>> rows,
            Map<String, String> taskProfile,
            LocalDate date,
            List<String> mergeIdentityHeaders) {
        String dc = ResultDispatchSchema.COL_DISPATCH_DATE;
        String qc = ResultDispatchSchema.COL_DISPATCH_QTY;
        double sum = 0;
        String ds = date.toString();
        for (Map<String, String> row : rows) {
            if (!matchesWideMergeIdentity(taskProfile, row, mergeIdentityHeaders)) {
                continue;
            }
            String rd = nz(row.get(dc));
            if (!ds.equals(rd) && !ds.equals(normalizeDateCell(rd))) {
                continue;
            }
            sum += ResultDispatchNormalizer.parseDouble(row.get(qc));
        }
        return sum;
    }

    /** Sum qty for (process, machine) on date (aggregated view). */
    public static double sumQuantityForProcessMachineDate(
            List<Map<String, String>> rows, String process, String machine, LocalDate date) {
        String dc = ResultDispatchSchema.COL_DISPATCH_DATE;
        String qc = ResultDispatchSchema.COL_DISPATCH_QTY;
        double sum = 0;
        String ds = date.toString();
        for (Map<String, String> row : rows) {
            if (!process.equals(nz(row.get(ResultDispatchSchema.COL_PROCESS)))) {
                continue;
            }
            if (!machine.equals(nz(row.get(ResultDispatchSchema.COL_MACHINE)))) {
                continue;
            }
            String rd = nz(row.get(dc));
            if (!ds.equals(rd) && !ds.equals(normalizeDateCell(rd))) {
                continue;
            }
            sum += ResultDispatchNormalizer.parseDouble(row.get(qc));
        }
        return sum;
    }

    /**
     * Whether {@code row} (canonical long-format) matches {@code profile} on all static columns (date/qty excluded).
     */
    public static boolean matchesTaskProfile(
            List<String> columns, Map<String, String> profile, Map<String, String> row) {
        return profileMatches(
                columns,
                profile,
                row,
                ResultDispatchSchema.COL_DISPATCH_DATE,
                ResultDispatchSchema.COL_DISPATCH_QTY);
    }

    /**
     * Same as {@link #matchesTaskProfile} but ignores {@link ResultDispatchSchema#COL_DISPATCH_TRIAL_ORDER} so callers
     * can align rows that still hold a previous trial-order value with a profile that already carries the next value.
     */
    public static boolean matchesTaskProfileExceptTrialOrder(
            List<String> columns, Map<String, String> profile, Map<String, String> row) {
        String dateCol = ResultDispatchSchema.COL_DISPATCH_DATE;
        String qtyCol = ResultDispatchSchema.COL_DISPATCH_QTY;
        String trialCol = ResultDispatchSchema.COL_DISPATCH_TRIAL_ORDER;
        for (String col : columns) {
            if (col.equals(dateCol) || col.equals(qtyCol) || col.equals(trialCol)) {
                continue;
            }
            if (!nz(profile.get(col)).equals(nz(row.get(col)))) {
                return false;
            }
        }
        return true;
    }

    static boolean profileMatches(
            List<String> columns,
            Map<String, String> profile,
            Map<String, String> row,
            String dateCol,
            String qtyCol) {
        for (String col : columns) {
            if (col.equals(dateCol) || col.equals(qtyCol)) {
                continue;
            }
            if (!nz(profile.get(col)).equals(nz(row.get(col)))) {
                return false;
            }
        }
        return true;
    }

    /** Replace allocations for profile+date with a single row of qty (0 removes). */
    public static void upsertAllocation(
            List<String> columns,
            List<Map<String, String>> rows,
            Map<String, String> taskProfile,
            LocalDate date,
            double qty) {
        String dc = ResultDispatchSchema.COL_DISPATCH_DATE;
        String qc = ResultDispatchSchema.COL_DISPATCH_QTY;
        String ds = date.toString();
        List<Map<String, String>> keep = new ArrayList<>();
        for (Map<String, String> row : rows) {
            String rd = nz(row.get(dc));
            boolean sameDate = ds.equals(rd) || ds.equals(normalizeDateCell(rd));
            if (sameDate && profileMatches(columns, taskProfile, row, dc, qc)) {
                continue;
            }
            keep.add(row);
        }
        rows.clear();
        rows.addAll(keep);
        if (qty > 1e-9) {
            LinkedHashMap<String, String> neo = new LinkedHashMap<>();
            for (String col : columns) {
                if (col.equals(dc)) {
                    neo.put(col, ds);
                } else if (col.equals(qc)) {
                    neo.put(col, ResultDispatchNormalizer.formatQty(qty));
                } else {
                    neo.put(col, nz(taskProfile.get(col)));
                }
            }
            rows.add(neo);
        }
        ResultDispatchNormalizer.normalizeInPlace(columns, rows);
    }

    /**
     * {@link #upsertAllocation} のワイド版: 同一タスク判定に {@link #matchesWideMergeIdentity} を使う。
     */
    public static void upsertAllocationForWideMerge(
            List<String> columns,
            List<Map<String, String>> rows,
            Map<String, String> taskProfile,
            LocalDate date,
            double qty,
            List<String> mergeIdentityHeaders) {
        String dc = ResultDispatchSchema.COL_DISPATCH_DATE;
        String qc = ResultDispatchSchema.COL_DISPATCH_QTY;
        String ds = date.toString();
        List<Map<String, String>> keep = new ArrayList<>();
        for (Map<String, String> row : rows) {
            String rd = nz(row.get(dc));
            boolean sameDate = ds.equals(rd) || ds.equals(normalizeDateCell(rd));
            if (sameDate && matchesWideMergeIdentity(taskProfile, row, mergeIdentityHeaders)) {
                continue;
            }
            keep.add(row);
        }
        rows.clear();
        rows.addAll(keep);
        if (qty > 1e-9) {
            LinkedHashMap<String, String> neo = new LinkedHashMap<>();
            for (String col : columns) {
                if (col.equals(dc)) {
                    neo.put(col, ds);
                } else if (col.equals(qc)) {
                    neo.put(col, ResultDispatchNormalizer.formatQty(qty));
                } else {
                    neo.put(col, nz(taskProfile.get(col)));
                }
            }
            rows.add(neo);
        }
        ResultDispatchNormalizer.normalizeInPlace(columns, rows);
    }

    /**
     * Scale rows matching process+machine+date so sum becomes {@code newTotal}. Distributes proportionally; if old sum is 0
     * assigns all to first matching profile row or creates minimal row from existing row shape.
     */
    public static void scaleProcessMachineDateToTotal(
            List<String> columns,
            List<Map<String, String>> rows,
            String process,
            String machine,
            LocalDate date,
            double newTotal) {
        String dc = ResultDispatchSchema.COL_DISPATCH_DATE;
        String qc = ResultDispatchSchema.COL_DISPATCH_QTY;
        String ds = date.toString();
        List<Map<String, String>> hits = new ArrayList<>();
        double oldSum = 0;
        for (Map<String, String> row : rows) {
            if (!process.equals(nz(row.get(ResultDispatchSchema.COL_PROCESS)))) {
                continue;
            }
            if (!machine.equals(nz(row.get(ResultDispatchSchema.COL_MACHINE)))) {
                continue;
            }
            String rd = nz(row.get(dc));
            if (!ds.equals(rd) && !ds.equals(normalizeDateCell(rd))) {
                continue;
            }
            hits.add(row);
            oldSum += ResultDispatchNormalizer.parseDouble(row.get(qc));
        }
        if (hits.isEmpty() && newTotal > 1e-9) {
            Map<String, String> seed =
                    rows.isEmpty() ? templateEmptyRow(columns) : new LinkedHashMap<>(rows.getFirst());
            seed.put(ResultDispatchSchema.COL_PROCESS, process);
            seed.put(ResultDispatchSchema.COL_MACHINE, machine);
            seed.put(dc, ds);
            seed.put(qc, ResultDispatchNormalizer.formatQty(newTotal));
            rows.add(seed);
            ResultDispatchNormalizer.normalizeInPlace(columns, rows);
            return;
        }
        if (hits.isEmpty()) {
            return;
        }
        if (oldSum <= 1e-9) {
            Map<String, String> first = hits.getFirst();
            first.put(qc, ResultDispatchNormalizer.formatQty(newTotal));
            for (int i = 1; i < hits.size(); i++) {
                rows.remove(hits.get(i));
            }
            ResultDispatchNormalizer.normalizeInPlace(columns, rows);
            return;
        }
        double factor = newTotal / oldSum;
        for (Map<String, String> h : hits) {
            double q = ResultDispatchNormalizer.parseDouble(h.get(qc));
            h.put(qc, ResultDispatchNormalizer.formatQty(q * factor));
        }
        ResultDispatchNormalizer.normalizeInPlace(columns, rows);
    }

    private static Map<String, String> templateEmptyRow(List<String> columns) {
        LinkedHashMap<String, String> m = new LinkedHashMap<>();
        for (String c : columns) {
            m.put(c, "");
        }
        return m;
    }

    static String normalizeDateCell(String raw) {
        LocalDate d = parseIsoDate(raw);
        return d != null ? d.toString() : nz(raw);
    }

    /**
     * Accepts {@code yyyy-MM-dd} (and common Excel/Japan display forms {@code yyyy/MM/dd}, {@code yyyy.MM.dd}).
     * Python 側の {@code _norm_ymd} はスラッシュ区切りのため、配台試行後の JSON 再読込でも日付軸と突き合わせできるよう、
     * 区切り文字の揺れ（{@code -} / {@code /} / {@code .}）に耐える。
     */
    public static LocalDate parseIsoDate(String raw) {
        if (raw == null || raw.isBlank()) {
            return null;
        }
        String t = raw.trim();
        if (t.length() >= 10 && t.charAt(4) == '-' && t.charAt(7) == '-') {
            try {
                return LocalDate.parse(t.substring(0, 10));
            } catch (DateTimeParseException e) {
                return null;
            }
        }
        if (t.length() >= 10) {
            char s1 = t.charAt(4);
            char s2 = t.charAt(7);
            if (s1 == s2 && (s1 == '/' || s1 == '.')) {
                try {
                    int y = Integer.parseInt(t.substring(0, 4));
                    int mo = Integer.parseInt(t.substring(5, 7));
                    int d = Integer.parseInt(t.substring(8, 10));
                    return LocalDate.of(y, mo, d);
                } catch (Exception e) {
                    return null;
                }
            }
        }
        return null;
    }

    /**
     * 同一工程・機械に紐づく行の「加工内容」を（空でないものだけ）抽出し、表示用に連結する。
     *
     * <p>列セットに「加工内容」が無い場合は空文字を返す。
     */
    public static String processingContentSummaryForProcessMachine(
            List<String> columns,
            List<Map<String, String>> rows,
            String process,
            String machine) {
        final String pc = "加工内容";
        if (columns == null || !columns.contains(pc)) {
            return "";
        }
        LinkedHashSet<String> distinct = new LinkedHashSet<>();
        for (Map<String, String> row : rows) {
            if (!nz(row.get(ResultDispatchSchema.COL_PROCESS)).equals(nz(process))) {
                continue;
            }
            if (!nz(row.get(ResultDispatchSchema.COL_MACHINE)).equals(nz(machine))) {
                continue;
            }
            String v = nz(row.get(pc));
            if (!v.isEmpty()) {
                distinct.add(v);
            }
        }
        if (distinct.isEmpty()) {
            return "";
        }
        return String.join(" / ", distinct);
    }

    public static List<Map.Entry<String, String>> sortedProcessMachineKeys(List<Map<String, String>> rows) {
        record Pm(String p, String m) {}
        Set<Pm> set = new TreeSet<>(Comparator.comparing(Pm::p).thenComparing(Pm::m));
        for (Map<String, String> row : rows) {
            String p = nz(row.get(ResultDispatchSchema.COL_PROCESS));
            String m = nz(row.get(ResultDispatchSchema.COL_MACHINE));
            if (!p.isEmpty() || !m.isEmpty()) {
                set.add(new Pm(p, m));
            }
        }
        List<Map.Entry<String, String>> out = new ArrayList<>();
        for (Pm pm : set) {
            out.add(Map.entry(pm.p, pm.m));
        }
        return out;
    }

    static String nz(String s) {
        return s != null ? s : "";
    }
}
