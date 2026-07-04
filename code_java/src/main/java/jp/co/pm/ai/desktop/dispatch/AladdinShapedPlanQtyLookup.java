package jp.co.pm.ai.desktop.dispatch;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Pattern;

import jp.co.pm.ai.desktop.io.JsonTableIo;

/**
 * アラジン加工計画（shaped JSON / 表）から {@code 機械名×依頼NO×日付×工程} の計画数量ルックアップを構築する。
 */
public final class AladdinShapedPlanQtyLookup {

    /** 原本転記・計画確認タブのアラジン計画日別列数。 */
    public static final int PIPELINE_CHECK_PLAN_DAY_COLUMNS = 7;

    /** 日付列ヘッダ: {@code yyyy/MM/dd} */
    private static final Pattern ALADDIN_DATE_COL = Pattern.compile("\\d{4}/\\d{2}/\\d{2}");

    private static final String COL_MK_NAME = "機械名";
    private static final String COL_TID = "依頼NO";
    private static final String COL_PROCESS = "工程名";

    private AladdinShapedPlanQtyLookup() {}

    public record ShapedTable(List<String> headers, List<List<String>> rows) {}

    /**
     * shaped 表ヘッダから日付列（{@code yyyy/MM/dd}）を昇順で抽出する。
     * {@code maxCount} を超える分は切り捨てる。
     */
    public static List<String> extractSortedDateColumnHeaders(List<String> headers, int maxCount) {
        if (headers == null || headers.isEmpty() || maxCount <= 0) {
            return List.of();
        }
        List<String> dates = new ArrayList<>();
        for (String h : headers) {
            if (h != null && ALADDIN_DATE_COL.matcher(h).matches()) {
                dates.add(h);
            }
        }
        dates.sort(String::compareTo);
        if (dates.size() <= maxCount) {
            return List.copyOf(dates);
        }
        return List.copyOf(dates.subList(0, maxCount));
    }

    /** 日付列見出しを表向け短ラベル（{@code M/d}）にする。 */
    public static String shortPlanDateColumnLabel(String dateYmd) {
        if (dateYmd == null || dateYmd.isBlank()) {
            return "計画日";
        }
        LocalDate d = parsePlanDateColumn(dateYmd);
        if (d != null) {
            return d.getMonthValue() + "/" + d.getDayOfMonth();
        }
        if (dateYmd.length() >= 10) {
            String tail = dateYmd.substring(5).replace("-", "/");
            if (tail.startsWith("0")) {
                tail = tail.substring(1);
            }
            int slash = tail.indexOf('/');
            if (slash >= 0 && slash + 1 < tail.length() && tail.charAt(slash + 1) == '0') {
                tail = tail.substring(0, slash + 1) + tail.substring(slash + 2);
            }
            return tail;
        }
        return dateYmd;
    }

    /** 原本転記・計画確認タブの計画日スロット列見出し（①〜⑳）。 */
    public static String circledSlotColumnLabel(int indexZeroBased) {
        if (indexZeroBased >= 0 && indexZeroBased < 20) {
            return String.valueOf((char) ('\u2460' + indexZeroBased));
        }
        return String.valueOf(indexZeroBased + 1);
    }

    /** 計画日セル表示（{@code M/d} + 半角スペース + {@code m} 付き数量）。数量なしは空文字。 */
    public static String formatPlanDateMetersCell(String dateYmd, double meters) {
        if (Math.abs(meters) <= 1e-12) {
            return "";
        }
        return shortPlanDateColumnLabel(dateYmd) + " " + formatPlanMeters(meters) + "m";
    }

    /**
     * 依頼の計画エントリを {@code dateHeaders} の各日付スロットに集計（同一日・複数機械は m 合算）。
     * 返却リスト長は {@code slotCount}（不足スロットは空文字）。各セルは {@code M/d Nm} 形式。
     */
    public static List<String> aggregatePlanMetersByDateSlots(
            List<PlanEntry> entries, List<String> dateHeaders, int slotCount) {
        if (slotCount <= 0) {
            return List.of();
        }
        Map<String, Double> sumByDate = new LinkedHashMap<>();
        if (entries != null) {
            for (PlanEntry e : entries) {
                String key = normaliseDateStr(e.dateYmd());
                if (key == null) {
                    key = e.dateYmd();
                }
                sumByDate.merge(key, e.planMeters(), Double::sum);
            }
        }
        List<String> out = new ArrayList<>(slotCount);
        for (int i = 0; i < slotCount; i++) {
            if (dateHeaders != null && i < dateHeaders.size()) {
                String header = dateHeaders.get(i);
                String key = normaliseDateStr(header);
                if (key == null) {
                    key = header;
                }
                Double qty = sumByDate.get(key);
                out.add(
                        qty != null && Math.abs(qty) > 1e-12
                                ? formatPlanDateMetersCell(header, qty)
                                : "");
            } else {
                out.add("");
            }
        }
        return List.copyOf(out);
    }

    private static String formatPlanMeters(double m) {
        if (Math.abs(m - Math.rint(m)) < 1e-9) {
            return String.valueOf((long) Math.rint(m));
        }
        return String.valueOf(m);
    }

    /** shaped JSON があれば読み込む。失敗・未存在時は空表。 */
    public static ShapedTable loadShapedTable(Path shapedJsonPath) {
        if (shapedJsonPath == null || !Files.isRegularFile(shapedJsonPath)) {
            return new ShapedTable(List.of(), List.of());
        }
        try {
            JsonTableIo.ArrayTable t = JsonTableIo.loadArrayTable(shapedJsonPath);
            return new ShapedTable(t.columns(), t.rows());
        } catch (Exception ex) {
            return new ShapedTable(List.of(), List.of());
        }
    }

    /**
     * Key: {@code normalizedMk -> tid -> yyyy/MM/dd -> processKey -> qty}. {@code processKey} は
     * {@link #normalizeProcessNameForRuleMatch}（工程名列があるとき）。無いときは {@code ""}。
     */
    public static Map<String, Map<String, Map<String, Map<String, Double>>>> buildLookup(
            List<String> headers, List<List<String>> rows) {
        int mkIdx = colIdx(headers, COL_MK_NAME);
        int tidIdx = colIdx(headers, COL_TID);
        int procIdx = colIdx(headers, COL_PROCESS);
        if (mkIdx < 0 || tidIdx < 0) {
            return Map.of();
        }
        Map<Integer, String> dateCols = new LinkedHashMap<>();
        for (int i = 0; i < headers.size(); i++) {
            String h = headers.get(i);
            if (h != null && ALADDIN_DATE_COL.matcher(h).matches()) {
                dateCols.put(i, h);
            }
        }
        if (dateCols.isEmpty()) {
            return Map.of();
        }
        Map<String, Map<String, Map<String, Map<String, Double>>>> result = new LinkedHashMap<>();
        for (List<String> row : rows) {
            String mk = normalizeEquipmentMatchKey(cellAt(row, mkIdx));
            String tid = cellAt(row, tidIdx).strip();
            if (mk.isEmpty() || tid.isEmpty()) {
                continue;
            }
            String procKey = "";
            if (procIdx >= 0) {
                procKey = normalizeProcessNameForRuleMatch(cellAt(row, procIdx));
            }
            for (Map.Entry<Integer, String> e : dateCols.entrySet()) {
                String dsRaw = e.getValue();
                String dsKey = normaliseDateStr(dsRaw);
                if (dsKey == null) {
                    dsKey = dsRaw;
                }
                double qty = parseCellDouble(cellAt(row, e.getKey()));
                if (Math.abs(qty) > 1e-12) {
                    result.computeIfAbsent(mk, k -> new LinkedHashMap<>())
                            .computeIfAbsent(tid, k -> new LinkedHashMap<>())
                            .computeIfAbsent(dsKey, k -> new LinkedHashMap<>())
                            .merge(procKey, qty, Double::sum);
                }
            }
        }
        return result;
    }

    public static double lookup(
            Map<String, Map<String, Map<String, Map<String, Double>>>> lookup,
            String mk,
            String tid,
            String dateStr,
            String processRaw) {
        Map<String, Map<String, Map<String, Double>>> byTid = lookup.get(normalizeEquipmentMatchKey(mk));
        if (byTid == null) {
            return 0.0;
        }
        Map<String, Map<String, Double>> byDate = byTid.get(tid != null ? tid.strip() : "");
        if (byDate == null) {
            return 0.0;
        }
        Map<String, Double> byProc = byDate.get(dateStr);
        if (byProc == null) {
            return 0.0;
        }
        if (byProc.size() == 1 && byProc.containsKey("")) {
            Double v = byProc.get("");
            return v != null ? v : 0.0;
        }
        String pk = normalizeProcessNameForRuleMatch(processRaw);
        if (pk.isEmpty()) {
            return 0.0;
        }
        Double v = byProc.get(pk);
        return v != null ? v : 0.0;
    }

    /**
     * ルックアップ内の {@code 機械名×依頼NO} に正の計画数量がある日付列（{@code yyyy/MM/dd}）を暦日に変換して返す。
     */
    public static List<LocalDate> distinctPlanDatesFor(
            Map<String, Map<String, Map<String, Map<String, Double>>>> lookup,
            String machineName,
            String taskId) {
        if (lookup == null || lookup.isEmpty()) {
            return List.of();
        }
        Map<String, Map<String, Map<String, Double>>> byTid =
                lookup.get(normalizeEquipmentMatchKey(machineName));
        if (byTid == null) {
            return List.of();
        }
        Map<String, Map<String, Double>> byDate = byTid.get(taskId != null ? taskId.strip() : "");
        if (byDate == null || byDate.isEmpty()) {
            return List.of();
        }
        Set<LocalDate> out = new LinkedHashSet<>();
        for (String dsKey : byDate.keySet()) {
            LocalDate d = parsePlanDateColumn(dsKey);
            if (d != null) {
                out.add(d);
            }
        }
        return new ArrayList<>(out);
    }

    /** 依頼NO に一致する全計画エントリ（機械名×工程名×日付×m）を収集する。 */
    public record PlanEntry(String machineName, String processName, String dateYmd, double planMeters) {}

    public static List<PlanEntry> collectEntriesForTaskId(
            Map<String, Map<String, Map<String, Map<String, Double>>>> lookup, String taskId) {
        if (lookup == null || lookup.isEmpty() || taskId == null || taskId.isBlank()) {
            return List.of();
        }
        String tid = taskId.strip();
        List<PlanEntry> out = new ArrayList<>();
        for (Map.Entry<String, Map<String, Map<String, Map<String, Double>>>> mkEntry :
                lookup.entrySet()) {
            Map<String, Map<String, Map<String, Double>>> byTid = mkEntry.getValue();
            Map<String, Map<String, Double>> byDate = byTid.get(tid);
            if (byDate == null || byDate.isEmpty()) {
                continue;
            }
            String machine = mkEntry.getKey();
            for (Map.Entry<String, Map<String, Double>> dateEntry : byDate.entrySet()) {
                String dateYmd = dateEntry.getKey();
                for (Map.Entry<String, Double> procEntry : dateEntry.getValue().entrySet()) {
                    double qty = procEntry.getValue() != null ? procEntry.getValue() : 0.0;
                    if (Math.abs(qty) > 1e-12) {
                        out.add(new PlanEntry(machine, procEntry.getKey(), dateYmd, qty));
                    }
                }
            }
        }
        out.sort(
                (a, b) -> {
                    int c = a.dateYmd().compareTo(b.dateYmd());
                    if (c != 0) {
                        return c;
                    }
                    c = a.machineName().compareTo(b.machineName());
                    if (c != 0) {
                        return c;
                    }
                    return a.processName().compareTo(b.processName());
                });
        return List.copyOf(out);
    }

    /** shaped 表から依頼NO の計画エントリを収集（表示用の機械名・工程名を保持）。 */
    public static List<PlanEntry> collectEntriesForTaskIdFromTable(
            List<String> headers, List<List<String>> rows, String taskId) {
        if (headers == null
                || rows == null
                || taskId == null
                || taskId.isBlank()) {
            return List.of();
        }
        int mkIdx = colIdx(headers, COL_MK_NAME);
        int tidIdx = colIdx(headers, COL_TID);
        int procIdx = colIdx(headers, COL_PROCESS);
        if (mkIdx < 0 || tidIdx < 0) {
            return List.of();
        }
        String tid = taskId.strip();
        Map<Integer, String> dateCols = new LinkedHashMap<>();
        for (int i = 0; i < headers.size(); i++) {
            String h = headers.get(i);
            if (h != null && ALADDIN_DATE_COL.matcher(h).matches()) {
                dateCols.put(i, h);
            }
        }
        if (dateCols.isEmpty()) {
            return List.of();
        }
        List<PlanEntry> out = new ArrayList<>();
        for (List<String> row : rows) {
            if (!tid.equals(cellAt(row, tidIdx).strip())) {
                continue;
            }
            String machine = cellAt(row, mkIdx).strip();
            String process = procIdx >= 0 ? cellAt(row, procIdx).strip() : "";
            for (Map.Entry<Integer, String> e : dateCols.entrySet()) {
                double qty = parseCellDouble(cellAt(row, e.getKey()));
                if (Math.abs(qty) > 1e-12) {
                    String dsKey = normaliseDateStr(e.getValue());
                    out.add(
                            new PlanEntry(
                                    machine,
                                    process,
                                    dsKey != null ? dsKey : e.getValue(),
                                    qty));
                }
            }
        }
        out.sort(
                (a, b) -> {
                    int c = a.dateYmd().compareTo(b.dateYmd());
                    if (c != 0) {
                        return c;
                    }
                    c = a.machineName().compareTo(b.machineName());
                    if (c != 0) {
                        return c;
                    }
                    return a.processName().compareTo(b.processName());
                });
        return List.copyOf(out);
    }

    /** {@code yyyy/MM/dd} または {@code yyyy-MM-dd} 形式の日付列キーを {@link LocalDate} に変換。 */
    public static LocalDate parsePlanDateColumn(String raw) {
        String n = normaliseDateStr(raw);
        if (n == null) {
            return null;
        }
        try {
            return LocalDate.parse(n, DateTimeFormatter.ofPattern("yyyy/MM/dd"));
        } catch (DateTimeParseException ex) {
            return null;
        }
    }

    /** Mirrors Python {@code _normalize_process_name_for_rule_match} (NFKC + remove spaces). */
    public static String normalizeProcessNameForRuleMatch(String raw) {
        if (raw == null) {
            return "";
        }
        String t = java.text.Normalizer.normalize(raw.strip(), java.text.Normalizer.Form.NFKC);
        return t.replaceAll("[\\s　]+", "");
    }

    static String normaliseDateStr(String raw) {
        if (raw == null || raw.isBlank()) {
            return null;
        }
        String s = raw.strip();
        if (s.length() >= 10 && s.charAt(4) == '-' && s.charAt(7) == '-') {
            return s.substring(0, 10).replace('-', '/');
        }
        if (s.length() == 10 && s.charAt(4) == '/' && s.charAt(7) == '/') {
            return s;
        }
        try {
            String[] parts = s.split("[/\\-]");
            if (parts.length == 3) {
                int y = Integer.parseInt(parts[0].strip());
                int mo = Integer.parseInt(parts[1].strip());
                int d = Integer.parseInt(parts[2].strip());
                return String.format("%04d/%02d/%02d", y, mo, d);
            }
        } catch (NumberFormatException ignored) {
            // fall through
        }
        return null;
    }

    private static String normalizeEquipmentMatchKey(String val) {
        if (val == null || val.isBlank()) {
            return "";
        }
        String t = java.text.Normalizer.normalize(val, java.text.Normalizer.Form.NFKC);
        t = t.replace('\u00a0', ' ').replace('\u3000', ' ');
        t = t.replaceAll("[\u200b\u200c\u200d\ufeff]", "");
        return t.replaceAll("\\s+", " ").strip();
    }

    private static double parseCellDouble(String s) {
        if (s == null || s.isBlank()) {
            return 0.0;
        }
        try {
            return Double.parseDouble(s.strip());
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }

    private static int colIdx(List<String> headers, String title) {
        for (int i = 0; i < headers.size(); i++) {
            if (title.equals(headers.get(i))) {
                return i;
            }
        }
        return -1;
    }

    private static String cellAt(List<String> row, int idx) {
        return (idx >= 0 && idx < row.size() && row.get(idx) != null) ? row.get(idx) : "";
    }
}
