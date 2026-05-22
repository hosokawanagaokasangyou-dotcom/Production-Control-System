package jp.co.pm.ai.desktop.dispatch;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import jp.co.pm.ai.desktop.io.JsonTableIo;

/**
 * アラジン加工計画（shaped JSON / 表）から {@code 機械名×依頼NO×日付×工程} の計画数量ルックアップを構築する。
 */
public final class AladdinShapedPlanQtyLookup {

    /** 日付列ヘッダ: {@code yyyy/MM/dd} */
    private static final Pattern ALADDIN_DATE_COL = Pattern.compile("\\d{4}/\\d{2}/\\d{2}");

    private static final String COL_MK_NAME = "機械名";
    private static final String COL_TID = "依頼NO";
    private static final String COL_PROCESS = "工程名";

    private AladdinShapedPlanQtyLookup() {}

    public record ShapedTable(List<String> headers, List<List<String>> rows) {}

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
