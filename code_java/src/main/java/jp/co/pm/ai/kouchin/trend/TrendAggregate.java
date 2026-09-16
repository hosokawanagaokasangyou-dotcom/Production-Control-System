package jp.co.pm.ai.kouchin.trend;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import jp.co.pm.ai.kouchin.verify.YearMonthKey;

/**
 * 工程グループ化と工場横断・前月差（振替疑い）。
 */
public final class TrendAggregate {

    private static final List<GroupRule> GROUP_RULES =
            List.of(
                    new GroupRule("スライス", List.of("スライス1", "スライス3", "スライス１", "スライス３", "スライス")),
                    new GroupRule("スリット", List.of("スリット4 ゴールド", "スリット4ゴールド", "スリット")),
                    new GroupRule("LAC(EC)", List.of("LAC(EC)", "ＬＡＣ(EC)", "LAC（EC）")),
                    new GroupRule("SEC(SS)", List.of("SEC(SS)", "SEC (SS)", "SEC（SS）")),
                    new GroupRule("SEC(EC)", List.of("SEC(EC)", "SEC (EC)", "SEC（EC）")));

    private record GroupRule(String group, List<String> aliases) {}

    private TrendAggregate() {}

    public static String processGroup(String name) {
        String n = TrendText.processName(name);
        if (n.isEmpty()) {
            return n;
        }
        for (GroupRule rule : GROUP_RULES) {
            for (String alias : rule.aliases) {
                if (n.equals(alias) || n.startsWith(alias)) {
                    if ("スライス".equals(rule.group) && n.startsWith("スライス")) {
                        return "スライス";
                    }
                    if ("スリット".equals(rule.group) && n.startsWith("スリット")) {
                        return "スリット";
                    }
                    if (n.equals(alias)) {
                        return rule.group;
                    }
                }
            }
        }
        for (GroupRule rule : GROUP_RULES) {
            for (String alias : rule.aliases) {
                if (n.equals(alias)) {
                    return rule.group;
                }
            }
        }
        if (n.startsWith("スライス")) {
            return "スライス";
        }
        if (n.startsWith("スリット")) {
            return "スリット";
        }
        return n;
    }

    public static TrendCrossMatrix buildCrossKind(
            List<TrendMonthData> kokubu,
            List<TrendMonthData> konan,
            List<YearMonthKey> months,
            String metric) {
        Map<YearMonthKey, TrendMonthData> kb = indexByYm(kokubu);
        Map<YearMonthKey, TrendMonthData> kn = indexByYm(konan);
        Set<String> processes = new HashSet<>();
        for (YearMonthKey ym : months) {
            processes.addAll(kindsOf(kb.get(ym)).keySet());
            processes.addAll(kindsOf(kn.get(ym)).keySet());
        }
        List<String> processesL = processes.stream().sorted().toList();
        Map<String, Map<String, Map<YearMonthKey, Double>>> values = new LinkedHashMap<>();
        values.put("国分", fillRaw(kb, processesL, months, metric));
        values.put("湖南", fillRaw(kn, processesL, months, metric));
        return new TrendCrossMatrix(processesL, List.copyOf(months), metric, values, false);
    }

    public static TrendCrossMatrix buildGroupedCross(
            List<TrendMonthData> kokubu,
            List<TrendMonthData> konan,
            List<YearMonthKey> months,
            String metric) {
        Map<YearMonthKey, TrendMonthData> kb = indexByYm(kokubu);
        Map<YearMonthKey, TrendMonthData> kn = indexByYm(konan);
        Set<String> groups = new HashSet<>();
        Map<String, Map<String, Map<YearMonthKey, Double>>> raw = new LinkedHashMap<>();
        raw.put("国分", new HashMap<>());
        raw.put("湖南", new HashMap<>());
        for (String factory : List.of("国分", "湖南")) {
            Map<YearMonthKey, TrendMonthData> src = "国分".equals(factory) ? kb : kn;
            for (YearMonthKey ym : months) {
                Map<String, Double> bucket = new HashMap<>();
                for (Map.Entry<String, TrendMetric> e : kindsOf(src.get(ym)).entrySet()) {
                    String g = processGroup(e.getKey());
                    if (g.isEmpty()) {
                        continue;
                    }
                    groups.add(g);
                    double add = e.getValue() == null ? 0.0 : e.getValue().value(metric);
                    bucket.merge(g, add, Double::sum);
                }
                for (Map.Entry<String, Double> e : bucket.entrySet()) {
                    raw.get(factory)
                            .computeIfAbsent(e.getKey(), k -> new HashMap<>())
                            .put(ym, e.getValue());
                }
            }
        }
        List<String> groupsL = groups.stream().sorted().toList();
        Map<String, Map<String, Map<YearMonthKey, Double>>> values = new LinkedHashMap<>();
        for (String factory : List.of("国分", "湖南")) {
            Map<YearMonthKey, TrendMonthData> src = "国分".equals(factory) ? kb : kn;
            Map<String, Map<YearMonthKey, Double>> fac = new LinkedHashMap<>();
            for (String g : groupsL) {
                Map<YearMonthKey, Double> ymMap = new HashMap<>();
                for (YearMonthKey ym : months) {
                    if (!src.containsKey(ym)) {
                        ymMap.put(ym, null);
                    } else {
                        ymMap.put(ym, raw.get(factory).getOrDefault(g, Map.of()).getOrDefault(ym, 0.0));
                    }
                }
                fac.put(g, ymMap);
            }
            values.put(factory, fac);
        }
        return new TrendCrossMatrix(groupsL, List.copyOf(months), metric, values, true);
    }

    public static String shiftBadge(Double dK, Double dH) {
        if (dK == null || dH == null) {
            return null;
        }
        if (dK == 0.0 && dH == 0.0) {
            return "変化なし";
        }
        if (dK * dH < 0) {
            double ratio = Math.min(Math.abs(dK), Math.abs(dH)) / Math.max(Math.abs(dK), Math.abs(dH));
            if (ratio >= 0.5) {
                return "強い相殺";
            }
            return "弱い相殺";
        }
        return "同方向";
    }

    public static List<TrendShiftRow> rankShiftSuspects(TrendCrossMatrix grouped, String metricLabel) {
        return rankShiftSuspects(grouped, metricLabel, 20);
    }

    public static List<TrendShiftRow> rankShiftSuspects(TrendCrossMatrix grouped, String metricLabel, int topN) {
        if (grouped == null || grouped.months().size() < 2) {
            return List.of();
        }
        YearMonthKey prev = grouped.months().get(grouped.months().size() - 2);
        YearMonthKey latest = grouped.months().get(grouped.months().size() - 1);
        Map<String, Map<String, Map<YearMonthKey, Double>>> values = grouped.values();
        List<TrendShiftRow> rows = new ArrayList<>();
        for (String proc : grouped.processes()) {
            Double k0 = cell(values, "国分", proc, prev);
            Double k1 = cell(values, "国分", proc, latest);
            Double h0 = cell(values, "湖南", proc, prev);
            Double h1 = cell(values, "湖南", proc, latest);
            if (k0 == null || k1 == null || h0 == null || h1 == null) {
                continue;
            }
            double dk = k1 - k0;
            double dh = h1 - h0;
            String badge = shiftBadge(dk, dh);
            double offset = dk * dh < 0 ? Math.min(Math.abs(dk), Math.abs(dh)) : 0.0;
            double combined = (k1 + h1) - (k0 + h0);
            rows.add(
                    new TrendShiftRow(
                            proc,
                            prev,
                            latest,
                            k0,
                            k1,
                            h0,
                            h1,
                            dk,
                            dh,
                            combined,
                            offset,
                            badge,
                            metricLabel == null ? "wage" : metricLabel));
        }
        rows.sort(
                Comparator.comparingInt((TrendShiftRow r) -> {
                            String b = r.badge();
                            return "強い相殺".equals(b) || "弱い相殺".equals(b) ? 0 : 1;
                        })
                        .thenComparing(Comparator.comparingDouble(TrendShiftRow::offset).reversed())
                        .thenComparing(
                                Comparator.comparingDouble(
                                                (TrendShiftRow r) -> Math.abs(r.dKokubu()) + Math.abs(r.dKonan()))
                                        .reversed()));
        if (rows.size() > topN) {
            return List.copyOf(rows.subList(0, topN));
        }
        return List.copyOf(rows);
    }

    public static List<String> pickFocusProcesses(
            TrendCrossMatrix groupedWage, TrendCrossMatrix groupedQty, int maxN) {
        List<String> focus = new ArrayList<>();
        if (groupedWage != null && groupedWage.processes().contains("スライス")) {
            focus.add("スライス");
        }
        for (TrendCrossMatrix src : new TrendCrossMatrix[] {groupedQty, groupedWage}) {
            if (src == null) {
                continue;
            }
            for (TrendShiftRow row : rankShiftSuspects(src, src.metric(), maxN)) {
                if (!focus.contains(row.process())) {
                    focus.add(row.process());
                }
                if (focus.size() >= maxN) {
                    return List.copyOf(focus);
                }
            }
        }
        if (groupedWage != null && !groupedWage.months().isEmpty()) {
            YearMonthKey latest = groupedWage.months().get(groupedWage.months().size() - 1);
            List<Map.Entry<Double, String>> sized = new ArrayList<>();
            for (String p : groupedWage.processes()) {
                double k = nz(cell(groupedWage.values(), "国分", p, latest));
                double h = nz(cell(groupedWage.values(), "湖南", p, latest));
                sized.add(Map.entry(k + h, p));
            }
            sized.sort(Comparator.<Map.Entry<Double, String>>comparingDouble(Map.Entry::getKey).reversed());
            for (Map.Entry<Double, String> e : sized) {
                if (!focus.contains(e.getValue())) {
                    focus.add(e.getValue());
                }
                if (focus.size() >= maxN) {
                    break;
                }
            }
        }
        return focus.size() > maxN ? List.copyOf(focus.subList(0, maxN)) : List.copyOf(focus);
    }

    public static Double factoryMonthTotal(List<TrendMonthData> monthsData, YearMonthKey ym, String metric) {
        TrendMonthData m = indexByYm(monthsData).get(ym);
        if (m == null) {
            return null;
        }
        double sum = 0.0;
        for (TrendMetric cell : kindsOf(m).values()) {
            sum += cell == null ? 0.0 : cell.value(metric);
        }
        return sum;
    }

    public static KindSeries buildKindSeries(List<TrendMonthData> monthsData, List<YearMonthKey> months, String metric) {
        Map<YearMonthKey, TrendMonthData> byYm = indexByYm(monthsData);
        Set<String> names = new HashSet<>();
        for (YearMonthKey ym : months) {
            names.addAll(kindsOf(byYm.get(ym)).keySet());
        }
        List<String> namesL = names.stream().sorted().toList();
        Map<String, Map<YearMonthKey, Double>> values = new LinkedHashMap<>();
        for (String name : namesL) {
            Map<YearMonthKey, Double> ymMap = new HashMap<>();
            for (YearMonthKey ym : months) {
                TrendMetric cell = kindsOf(byYm.get(ym)).get(name);
                ymMap.put(ym, cell == null ? null : cell.value(metric));
            }
            values.put(name, ymMap);
        }
        return new KindSeries(namesL, List.copyOf(months), metric, values);
    }

    public static SheetSeries buildSheetSeries(
            List<TrendMonthData> monthsData, List<YearMonthKey> months, String metric) {
        Map<YearMonthKey, TrendMonthData> byYm = indexByYm(monthsData);
        Set<String> names = new HashSet<>();
        for (YearMonthKey ym : months) {
            names.addAll(sheetsOf(byYm.get(ym)).keySet());
        }
        List<String> namesL = names.stream().sorted().toList();
        Map<String, Map<YearMonthKey, Double>> values = new LinkedHashMap<>();
        for (String name : namesL) {
            Map<YearMonthKey, Double> ymMap = new HashMap<>();
            for (YearMonthKey ym : months) {
                TrendMetric cell = sheetsOf(byYm.get(ym)).get(name);
                ymMap.put(ym, cell == null ? null : cell.value(metric));
            }
            values.put(name, ymMap);
        }
        return new SheetSeries(namesL, List.copyOf(months), metric, values);
    }

    public record KindSeries(
            List<String> processes,
            List<YearMonthKey> months,
            String metric,
            Map<String, Map<YearMonthKey, Double>> values) {}

    public record SheetSeries(
            List<String> sheetNames,
            List<YearMonthKey> months,
            String metric,
            Map<String, Map<YearMonthKey, Double>> values) {}

    public static YearMonthKey[] lastComparablePair(
            List<YearMonthKey> months, List<TrendMonthData> kokubu, List<TrendMonthData> konan) {
        Set<YearMonthKey> kb = new HashSet<>();
        for (TrendMonthData m : kokubu) {
            if (m.ym() != null && !kindsOf(m).isEmpty()) {
                kb.add(m.ym());
            }
        }
        Set<YearMonthKey> kn = new HashSet<>();
        for (TrendMonthData m : konan) {
            if (m.ym() != null && !kindsOf(m).isEmpty()) {
                kn.add(m.ym());
            }
        }
        List<YearMonthKey> both = new ArrayList<>();
        for (YearMonthKey ym : months) {
            if (kb.contains(ym) && kn.contains(ym)) {
                both.add(ym);
            }
        }
        if (both.size() < 2) {
            return new YearMonthKey[] {null, null};
        }
        return new YearMonthKey[] {both.get(both.size() - 2), both.get(both.size() - 1)};
    }

    private static Map<YearMonthKey, TrendMonthData> indexByYm(List<TrendMonthData> monthsData) {
        Map<YearMonthKey, TrendMonthData> map = new TreeMap<>();
        if (monthsData == null) {
            return map;
        }
        for (TrendMonthData m : monthsData) {
            if (m != null && m.ym() != null) {
                map.put(m.ym(), m);
            }
        }
        return map;
    }

    private static Map<String, TrendMetric> kindsOf(TrendMonthData m) {
        return m == null || m.kinds() == null ? Map.of() : m.kinds();
    }

    private static Map<String, TrendMetric> sheetsOf(TrendMonthData m) {
        return m == null || m.sheets() == null ? Map.of() : m.sheets();
    }

    private static Map<String, Map<YearMonthKey, Double>> fillRaw(
            Map<YearMonthKey, TrendMonthData> src,
            List<String> processes,
            List<YearMonthKey> months,
            String metric) {
        Map<String, Map<YearMonthKey, Double>> fac = new LinkedHashMap<>();
        for (String proc : processes) {
            Map<YearMonthKey, Double> ymMap = new HashMap<>();
            for (YearMonthKey ym : months) {
                TrendMetric cell = kindsOf(src.get(ym)).get(proc);
                ymMap.put(ym, cell == null ? null : cell.value(metric));
            }
            fac.put(proc, ymMap);
        }
        return fac;
    }

    private static Double cell(
            Map<String, Map<String, Map<YearMonthKey, Double>>> values,
            String factory,
            String proc,
            YearMonthKey ym) {
        Map<String, Map<YearMonthKey, Double>> fac = values.get(factory);
        if (fac == null) {
            return null;
        }
        Map<YearMonthKey, Double> procMap = fac.get(proc);
        if (procMap == null) {
            return null;
        }
        return procMap.get(ym);
    }

    private static double nz(Double v) {
        return v == null ? 0.0 : v;
    }
}
