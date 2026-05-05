package jp.co.pm.ai.desktop.dispatch;

import java.text.Normalizer;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Set;

/**
 * Compares a saved result-dispatch document with the post-trial JSON on disk.
 *
 * <p>Quantities: aggregated by (依頼NO, 機械名) — same total rule as Python {@code
 * _interactive_validate_dispatch_quantities} for UI trial. Per-calendar-day rows may differ when the
 * trial output splits multi-day work while the editor snapshot had one aggregate row.
 *
 * <p>Order: minimum {@link ResultDispatchSchema#COL_DISPATCH_TRIAL_ORDER} per (依頼NO, 工程名, 機械名),
 * aligned with {@code merge_interactive_result_dispatch_json_into_tasks_df}.
 */
public final class DispatchTrialConsistency {

    private static final double QTY_EPS = 1e-3;

    /** Quantity total key: same grouping as Python (tid, mach) sums. */
    public record QtyPair(String requestNo, String machineName) {}

    /** Order key aligned with Python {@code order_map (tid, proc, mach)}. */
    public record OrderTriple(String requestNo, String processName, String machineName) {}

    public record CheckResult(boolean consistent, List<String> detailLines) {}

    private DispatchTrialConsistency() {}

    public static CheckResult compareDocuments(ResultDispatchDocument before, ResultDispatchDocument after) {
        List<String> lines = new ArrayList<>();
        Map<QtyPair, Double> bQty = aggregateQuantitiesByPair(before);
        Map<QtyPair, Double> aQty = aggregateQuantitiesByPair(after);
        compareQtyPairMaps(bQty, aQty, lines);

        Map<OrderTriple, Integer> bOrd = minTrialOrders(before);
        Map<OrderTriple, Integer> aOrd = minTrialOrders(after);
        compareOrderMaps(bOrd, aOrd, lines);

        return new CheckResult(lines.isEmpty(), lines);
    }

    static Map<QtyPair, Double> aggregateQuantitiesByPair(ResultDispatchDocument doc) {
        Map<QtyPair, Double> out = new LinkedHashMap<>();
        if (doc == null) {
            return out;
        }
        String reqCol = "依頼NO";
        for (Map<String, String> row : doc.rows()) {
            String tid = normCell(row.get(reqCol));
            String mach = normCell(row.get(ResultDispatchSchema.COL_MACHINE));
            LocalDate dd = parseDispatchDate(row.get(ResultDispatchSchema.COL_DISPATCH_DATE));
            double q = ResultDispatchNormalizer.parseDouble(row.get(ResultDispatchSchema.COL_DISPATCH_QTY));
            if (dd != null && !tid.isEmpty() && !mach.isEmpty() && q > QTY_EPS) {
                QtyPair k = new QtyPair(tid, mach);
                out.merge(k, q, Double::sum);
            }
        }
        return out;
    }

    static Map<OrderTriple, Integer> minTrialOrders(ResultDispatchDocument doc) {
        Map<OrderTriple, Integer> out = new LinkedHashMap<>();
        if (doc == null) {
            return out;
        }
        String reqCol = "依頼NO";
        for (Map<String, String> row : doc.rows()) {
            String tid = normCell(row.get(reqCol));
            String proc = normCell(row.get(ResultDispatchSchema.COL_PROCESS));
            String mach = normCell(row.get(ResultDispatchSchema.COL_MACHINE));
            Integer o = parseOrderInt(row.get(ResultDispatchSchema.COL_DISPATCH_TRIAL_ORDER));
            if (o == null || tid.isEmpty()) {
                continue;
            }
            OrderTriple k = new OrderTriple(tid, proc, mach);
            out.merge(k, o, Math::min);
        }
        return out;
    }

    private static void compareQtyPairMaps(
            Map<QtyPair, Double> before, Map<QtyPair, Double> after, List<String> lines) {
        Set<QtyPair> keys = new LinkedHashSet<>();
        keys.addAll(before.keySet());
        keys.addAll(after.keySet());
        for (QtyPair k : keys) {
            double bv = before.getOrDefault(k, 0.0);
            double av = after.getOrDefault(k, 0.0);
            if (Math.abs(bv - av) > QTY_EPS) {
                lines.add(
                        "[数量・依頼NO×機械名合計] 依頼NO="
                                + k.requestNo()
                                + " 機械="
                                + k.machineName()
                                + " … 試行前合計="
                                + formatNum(bv)
                                + " / 試行後合計="
                                + formatNum(av));
            }
        }
    }

    private static void compareOrderMaps(
            Map<OrderTriple, Integer> before, Map<OrderTriple, Integer> after, List<String> lines) {
        Set<OrderTriple> keys = new LinkedHashSet<>();
        keys.addAll(before.keySet());
        keys.addAll(after.keySet());
        for (OrderTriple k : keys) {
            Integer bo = before.get(k);
            Integer ao = after.get(k);
            if (Objects.equals(bo, ao)) {
                continue;
            }
            lines.add(
                    "[配台試行順番] 依頼NO="
                            + k.requestNo()
                            + " 工程="
                            + k.processName()
                            + " 機械="
                            + k.machineName()
                            + " … 試行前最小順="
                            + (bo != null ? bo : "—")
                            + " / 試行後最小順="
                            + (ao != null ? ao : "—"));
        }
    }

    private static String formatNum(double v) {
        if (Double.isNaN(v) || Double.isInfinite(v)) {
            return "0";
        }
        if (Math.abs(v - Math.rint(v)) < 1e-9 && Math.abs(v) < 1e15) {
            return Long.toString((long) Math.rint(v));
        }
        return Objects.toString(v);
    }

    static String normCell(String raw) {
        if (raw == null || raw.isBlank()) {
            return "";
        }
        return Normalizer.normalize(raw.strip(), Normalizer.Form.NFKC).trim();
    }

    private static final DateTimeFormatter SLASH_DF =
            DateTimeFormatter.ofPattern("yyyy/M/d", Locale.JAPAN);

    static LocalDate parseDispatchDate(String raw) {
        if (raw == null || raw.isBlank()) {
            return null;
        }
        String s = normCell(raw);
        if (s.length() >= 10 && s.charAt(4) == '-' && s.charAt(7) == '-') {
            try {
                return LocalDate.parse(s.substring(0, 10));
            } catch (DateTimeParseException ignored) {
                return null;
            }
        }
        if (s.contains("/")) {
            try {
                return LocalDate.parse(s, SLASH_DF);
            } catch (DateTimeParseException e) {
                try {
                    String[] p = s.split("/");
                    if (p.length >= 3) {
                        int y = Integer.parseInt(p[0].trim());
                        int m = Integer.parseInt(p[1].trim());
                        int d = Integer.parseInt(p[2].trim());
                        return LocalDate.of(y, m, d);
                    }
                } catch (Exception ignored) {
                }
                return null;
            }
        }
        return null;
    }

    /**
     * Parses 配台試行順番; empty cell → {@code null} (Python skips), including {@code 0} as in
     * {@code merge_interactive_result_dispatch_json_into_tasks_df}.
     */
    private static Integer parseOrderInt(String raw) {
        if (raw == null || raw.isBlank()) {
            return null;
        }
        String t = normCell(raw).replace(",", "");
        try {
            return (int) Math.round(Double.parseDouble(t));
        } catch (NumberFormatException e) {
            return null;
        }
    }
}
