package jp.co.pm.ai.kouchin.trend;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import jp.co.pm.ai.kouchin.verify.YearMonthKey;

/**
 * 月次トレンドの集計結果。賃・量は千円／km（表示単位）。
 */
public record TrendModel(
        List<YearMonthKey> months,
        List<String> focusProcesses,
        TrendCrossMatrix wage,
        TrendCrossMatrix qty,
        List<TrendShiftRow> suspects,
        List<TrendMonthData> kokubu,
        List<TrendMonthData> konan,
        List<String> warnings) {

    public TrendModel {
        months = months == null ? List.of() : List.copyOf(months);
        focusProcesses = focusProcesses == null ? List.of() : List.copyOf(focusProcesses);
        suspects = suspects == null ? List.of() : List.copyOf(suspects);
        kokubu = kokubu == null ? List.of() : List.copyOf(kokubu);
        konan = konan == null ? List.of() : List.copyOf(konan);
        warnings = warnings == null ? List.of() : List.copyOf(warnings);
    }

    /**
     * 生値（円／m）から表示単位のモデルを組み立てる。
     */
    public static TrendModel assemble(
            List<YearMonthKey> months,
            List<TrendMonthData> kokubuRaw,
            List<TrendMonthData> konanRaw,
            List<String> warnings) {
        List<TrendMonthData> kokubu = toDisplayUnits(kokubuRaw);
        List<TrendMonthData> konan = toDisplayUnits(konanRaw);
        List<YearMonthKey> ms = months == null ? List.of() : months;
        TrendCrossMatrix gw = TrendAggregate.buildGroupedCross(kokubu, konan, ms, "wage");
        TrendCrossMatrix gq = TrendAggregate.buildGroupedCross(kokubu, konan, ms, "qty");
        List<String> focus = TrendAggregate.pickFocusProcesses(gw, gq, 6);
        List<TrendShiftRow> suspects = TrendAggregate.rankShiftSuspects(gq, "qty");
        if (suspects.isEmpty()) {
            suspects = TrendAggregate.rankShiftSuspects(gw, "wage");
        }
        return new TrendModel(ms, focus, gw, gq, suspects, kokubu, konan, warnings);
    }

    static List<TrendMonthData> toDisplayUnits(List<TrendMonthData> src) {
        List<TrendMonthData> out = new ArrayList<>();
        if (src == null) {
            return out;
        }
        for (TrendMonthData m : src) {
            Map<String, TrendMetric> kinds = new LinkedHashMap<>();
            for (var e : m.kinds().entrySet()) {
                kinds.put(e.getKey(), convert(e.getValue()));
            }
            Map<String, TrendMetric> sheets = new LinkedHashMap<>();
            for (var e : m.sheets().entrySet()) {
                sheets.put(e.getKey(), convert(e.getValue()));
            }
            out.add(new TrendMonthData(m.ym(), m.path(), kinds, sheets, m.warnings()));
        }
        return out;
    }

    private static TrendMetric convert(TrendMetric m) {
        if (m == null) {
            return new TrendMetric(0, 0);
        }
        return new TrendMetric(TrendText.roundDisp(m.wage() / 1000.0), TrendText.roundDisp(m.qty() / 1000.0));
    }
}
