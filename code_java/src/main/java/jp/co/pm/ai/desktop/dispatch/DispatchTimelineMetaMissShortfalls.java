package jp.co.pm.ai.desktop.dispatch;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 段階3試行後、計画数量があるのにタイムライン未割付（加工開始日時が空）の行を検出する。
 *
 * <p>Python {@code dispatch_trial_shortages.json} が空／古い場合の Java 側フォールバック。
 */
public final class DispatchTimelineMetaMissShortfalls {

    private static final double EPS = 1e-6;
    private static final DateTimeFormatter SLASH_DAY = DateTimeFormatter.ofPattern("uuuu/M/d");

    static final String NOTE =
            "計画暦日にタイムライン割付なし（加工開始日時が空）。"
                    + "段階2の時刻は段階3試行でクリアされています。";

    private DispatchTimelineMetaMissShortfalls() {}

    public static List<DispatchTrialShortages.DispatchQtyShortfallRow> detectFromDocument(
            ResultDispatchDocument doc) {
        if (doc == null) {
            return List.of();
        }
        return detectFromRows(doc.columns(), doc.rows());
    }

    static List<DispatchTrialShortages.DispatchQtyShortfallRow> detectFromRows(
            List<String> columns, List<Map<String, String>> rows) {
        if (rows == null || rows.isEmpty()) {
            return List.of();
        }
        Map<String, DispatchTrialShortages.DispatchQtyShortfallRow> byKey = new LinkedHashMap<>();
        for (Map<String, String> row : rows) {
            if (row == null) {
                continue;
            }
            double plan =
                    ResultDispatchNormalizer.parseDouble(
                            row.get(ResultDispatchSchema.COL_DISPATCH_QTY));
            if (plan <= EPS) {
                continue;
            }
            String start = nz(row.get("加工開始日時"));
            if (!start.isEmpty()) {
                continue;
            }
            String tid = nz(row.get("依頼NO"));
            if (tid.isEmpty()) {
                tid = nz(row.get("タスクID"));
            }
            String mach = nz(row.get(ResultDispatchSchema.COL_MACHINE));
            LocalDate dd = parseDispatchDay(row.get(ResultDispatchSchema.COL_DISPATCH_DATE));
            if (tid.isEmpty() || mach.isEmpty() || dd == null) {
                continue;
            }
            String dateIso = dd.toString();
            String key =
                    DispatchTrialShortages.wideShortfallKey(tid, mach, dateIso);
            byKey.putIfAbsent(
                    key,
                    new DispatchTrialShortages.DispatchQtyShortfallRow(
                            tid, mach, dateIso, plan, 0.0, plan, NOTE));
        }
        return List.copyOf(byKey.values());
    }

    private static LocalDate parseDispatchDay(String cell) {
        if (cell == null) {
            return null;
        }
        String t = cell.trim();
        if (t.isEmpty()) {
            return null;
        }
        if (t.contains("T")) {
            t = t.substring(0, t.indexOf('T'));
        }
        try {
            return LocalDate.parse(t);
        } catch (DateTimeParseException ignored) {
        }
        try {
            return LocalDate.parse(t, SLASH_DAY);
        } catch (DateTimeParseException ignored) {
        }
        return null;
    }

    private static String nz(String s) {
        return s != null ? s.strip() : "";
    }
}
