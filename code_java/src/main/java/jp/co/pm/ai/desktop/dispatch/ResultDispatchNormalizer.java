package jp.co.pm.ai.desktop.dispatch;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * Merges duplicate rows that share the same group key and {@link ResultDispatchSchema#COL_DISPATCH_DATE},
 * summing {@link ResultDispatchSchema#COL_DISPATCH_QTY}.
 */
public final class ResultDispatchNormalizer {

    private ResultDispatchNormalizer() {}

    public static void normalizeInPlace(List<String> columns, List<Map<String, String>> rows) {
        if (rows.isEmpty()) {
            return;
        }
        String dateCol = ResultDispatchSchema.COL_DISPATCH_DATE;
        String qtyCol = ResultDispatchSchema.COL_DISPATCH_QTY;
        Map<String, Map<String, String>> acc = new LinkedHashMap<>();
        for (Map<String, String> row : rows) {
            String gk = groupKey(columns, row, dateCol, qtyCol);
            String dk = nz(row.get(dateCol));
            String key = gk + "\u0000" + dk;
            double q = parseDouble(row.get(qtyCol));
            Map<String, String> existing = acc.get(key);
            if (existing == null) {
                Map<String, String> copy = new LinkedHashMap<>(row);
                copy.put(qtyCol, formatQty(q));
                acc.put(key, copy);
            } else {
                double prev = parseDouble(existing.get(qtyCol));
                existing.put(qtyCol, formatQty(prev + q));
            }
        }
        rows.clear();
        rows.addAll(acc.values());
    }

    static String groupKey(List<String> columns, Map<String, String> row, String dateCol, String qtyCol) {
        StringBuilder sb = new StringBuilder();
        for (String c : columns) {
            if (c.equals(dateCol) || c.equals(qtyCol)) {
                continue;
            }
            sb.append('|');
            sb.append(nz(row.get(c)));
        }
        return sb.toString();
    }

    static String nz(String s) {
        return s != null ? s : "";
    }

    public static double parseDouble(String s) {
        if (s == null || s.isBlank()) {
            return 0d;
        }
        try {
            return Double.parseDouble(s.trim().replace(",", ""));
        } catch (NumberFormatException e) {
            return 0d;
        }
    }

    public static String formatQty(double v) {
        if (Double.isNaN(v) || Double.isInfinite(v)) {
            return "0";
        }
        if (v == Math.rint(v) && Math.abs(v) < 1e15) {
            return Long.toString((long) v);
        }
        return Objects.toString(v);
    }
}
