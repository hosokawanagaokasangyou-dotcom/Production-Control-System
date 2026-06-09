package jp.co.pm.ai.desktop.io.actuals;

import java.text.Collator;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanMemberLookup;
import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchNormalizer;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchSchema;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.io.TaskInputSourceRawGridIo;
import jp.co.pm.ai.desktop.io.gantt.PersonNameHeuristics;

/**
 * 加工実績・アラジン予定・配台予定を機械名単位に集約する。
 */
public final class EquipmentStatusDashboardBuilder {

    private static final Collator MACHINE_COLLATOR = Collator.getInstance(Locale.JAPAN);

    private static final String COL_MACHINE = "機械名";
    private static final String COL_REQUEST = "依頼NO";
    private static final String COL_PROCESS = "工程名";
    private static final String COL_KAKOU_DATE = "加工日";
    private static final String COL_START_DT = "加工開始日時";
    private static final String COL_END_DT = "加工終了日時";
    private static final String COL_QTY_CONV = "換算数量";
    private static final String COL_ACTUAL_QTY = "実加工数";
    private static final String COL_CUM_DONE = "累積実績";
    private static final String COL_CUM_PCT = "累積完了率";
    private static final String COL_MEMBER = "メンバー名";

    /** 実績明細で担当者を探す列（表記揺れ）。 */
    private static final String[] ACTUAL_MEMBER_COLUMNS = {
        "メンバー名", "担当OP_指定", "担当OP指定", "担当OP", "担当者名", "OP名"
    };

    private static final Pattern PCT_SUFFIX = Pattern.compile("%\\s*$");
    private static final DateTimeFormatter DT_OUT =
            DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm");

    static {
        MACHINE_COLLATOR.setStrength(Collator.PRIMARY);
    }

    private EquipmentStatusDashboardBuilder() {}

    public static List<EquipmentMachineStatus> build(
            ActualsSnapshot actuals,
            AladdinSnapshot aladdin,
            DispatchSnapshot dispatch,
            LocalDate actualDate,
            LocalDate planDate) {
        return build(actuals, aladdin, dispatch, actualDate, planDate, LocalDate.now());
    }

    /** @param statusToday 稼働中判定の「当日」（テスト用に固定可能）。 */
    static List<EquipmentMachineStatus> build(
            ActualsSnapshot actuals,
            AladdinSnapshot aladdin,
            DispatchSnapshot dispatch,
            LocalDate actualDate,
            LocalDate planDate,
            LocalDate statusToday) {
        List<String> actHeaders = actuals != null ? actuals.headers() : List.of();
        List<List<String>> actRows = actuals != null ? actuals.rows() : List.of();
        List<String> alHeaders = aladdin != null ? aladdin.headers() : List.of();
        List<List<String>> alRows = aladdin != null ? aladdin.rows() : List.of();
        List<String> disHeaders = dispatch != null ? dispatch.headers() : List.of();
        List<List<String>> disRows = dispatch != null ? dispatch.rows() : List.of();

        String actualDateKey = formatDateKey(actualDate);
        String planDateKey = formatDateKey(planDate);

        Map<String, List<List<String>>> actualByMachine =
                groupActualRowsByMachine(actHeaders, actRows, actualDateKey);
        Set<String> machines = new LinkedHashSet<>();
        machines.addAll(actualByMachine.keySet());
        machines.addAll(allMachinesInTable(alHeaders, alRows, COL_MACHINE));
        machines.addAll(allMachinesInTable(disHeaders, disRows, ResultDispatchSchema.COL_MACHINE));

        List<String> sorted = new ArrayList<>(machines);
        sorted.sort(MACHINE_COLLATOR);

        List<EquipmentMachineStatus> out = new ArrayList<>(sorted.size());
        for (String machine : sorted) {
            List<List<String>> dayRows = actualByMachine.getOrDefault(machine, List.of());
            Optional<EquipmentMachineStatus.ActualTaskRow> task =
                    pickLatestActualTask(actHeaders, dayRows);
            if (task.isPresent()) {
                task =
                        Optional.of(
                                enrichActualTaskMember(
                                        task.get(),
                                        alHeaders,
                                        alRows,
                                        disHeaders,
                                        disRows,
                                        machine,
                                        actualDateKey,
                                        planDateKey));
                task =
                        applyDayCompletionPct(
                                task,
                                actHeaders,
                                dayRows,
                                alHeaders,
                                alRows,
                                machine,
                                actualDateKey);
            }
            EquipmentMachineStatus.Status status =
                    deriveStatus(
                            actHeaders,
                            dayRows,
                            task,
                            alHeaders,
                            alRows,
                            machine,
                            actualDate,
                            statusToday);
            List<EquipmentMachineStatus.PlanLine> alPlans =
                    collectAladdinPlans(alHeaders, alRows, machine, planDateKey);
            List<EquipmentMachineStatus.PlanLine> disPlans =
                    collectDispatchPlans(disHeaders, disRows, machine, planDateKey);
            out.add(
                    new EquipmentMachineStatus(
                            machine, status, task, alPlans, disPlans));
        }
        return List.copyOf(out);
    }

    /** {@link PlanInputTabularIo.TabularSheet} から実績スナップショットを生成。 */
    public static ActualsSnapshot actualsFrom(PlanInputTabularIo.TabularSheet sheet) {
        if (sheet == null) {
            return new ActualsSnapshot(List.of(), List.of());
        }
        PlanInputTabularIo.TabularSheet deduped =
                TaskInputSourceRawGridIo.applyProcessingActualsDedupeByQuadKey(sheet);
        return new ActualsSnapshot(
                copyList(deduped.headers()), copyMatrix(deduped.rows()));
    }

    /** ロット別明細の重複行を除いてからダッシュボード集計に渡す。 */
    public static ActualsSnapshot normalizeActualsSnapshot(ActualsSnapshot snapshot) {
        if (snapshot == null) {
            return new ActualsSnapshot(List.of(), List.of());
        }
        return actualsFrom(
                new PlanInputTabularIo.TabularSheet(snapshot.headers(), snapshot.rows()));
    }

    public static AladdinSnapshot aladdinFrom(PlanInputTabularIo.TabularSheet sheet) {
        if (sheet == null) {
            return new AladdinSnapshot(List.of(), List.of());
        }
        return new AladdinSnapshot(
                copyList(sheet.headers()), copyMatrix(sheet.rows()));
    }

    public static DispatchSnapshot dispatchFrom(PlanInputTabularIo.TabularSheet sheet) {
        if (sheet == null) {
            return new DispatchSnapshot(List.of(), List.of());
        }
        return new DispatchSnapshot(
                copyList(sheet.headers()), copyMatrix(sheet.rows()));
    }

    static double parseCompletionPct(List<String> headers, List<String> row) {
        int iCumPct = colIdx(headers, COL_CUM_PCT);
        if (iCumPct >= 0) {
            double p = parsePctCell(cellAt(row, iCumPct));
            if (Double.isFinite(p)) {
                return clampPct(p);
            }
        }
        int iCum = colIdx(headers, COL_CUM_DONE);
        int iConv = colIdx(headers, COL_QTY_CONV);
        if (iCum >= 0 && iConv >= 0) {
            double cum = parseDouble(cellAt(row, iCum));
            double conv = parseDouble(cellAt(row, iConv));
            if (conv > 1e-12) {
                return clampPct(cum / conv * 100.0);
            }
        }
        int iAct = colIdx(headers, COL_ACTUAL_QTY);
        if (iAct >= 0 && iConv >= 0) {
            double act = parseDouble(cellAt(row, iAct));
            double conv = parseDouble(cellAt(row, iConv));
            if (conv > 1e-12) {
                return clampPct(act / conv * 100.0);
            }
        }
        return 0.0;
    }

    static LocalDate rowActualDate(List<String> headers, List<String> row) {
        int iStart = colIdx(headers, COL_START_DT);
        if (iStart >= 0) {
            LocalDate d = parseDatePrefix(cellAt(row, iStart));
            if (d != null) {
                return d;
            }
        }
        int iDate = colIdx(headers, COL_KAKOU_DATE);
        if (iDate >= 0) {
            return parseDateKey(cellAt(row, iDate));
        }
        return null;
    }

    private static EquipmentMachineStatus.Status deriveStatus(
            List<String> actHeaders,
            List<List<String>> dayRows,
            Optional<EquipmentMachineStatus.ActualTaskRow> task,
            List<String> alHeaders,
            List<List<String>> alRows,
            String machine,
            LocalDate actualDate,
            LocalDate statusToday) {
        if (dayRows == null || dayRows.isEmpty() || task.isEmpty()) {
            return EquipmentMachineStatus.Status.STOPPED;
        }
        if (actualDate != null
                && statusToday != null
                && actualDate.equals(statusToday)) {
            double actualQty = sumActualQtyM(actHeaders, dayRows);
            double aladdinPlan =
                    sumAladdinPlanQtyM(
                            alHeaders, alRows, machine, formatDateKey(actualDate));
            if (actualQty <= 1e-12) {
                return EquipmentMachineStatus.Status.STOPPED;
            }
            if (actualQty + 1e-12 < aladdinPlan) {
                return EquipmentMachineStatus.Status.RUNNING;
            }
            return EquipmentMachineStatus.Status.COMPLETED;
        }
        double pct = task.get().completionPct();
        if (pct >= 99.999) {
            return EquipmentMachineStatus.Status.COMPLETED;
        }
        return EquipmentMachineStatus.Status.RUNNING;
    }

    /** 実績表示日の機械別実績合計（m）。 */
    static double sumActualQtyM(List<String> headers, List<List<String>> dayRows) {
        if (dayRows == null || dayRows.isEmpty()) {
            return 0.0;
        }
        double sum = 0.0;
        for (List<String> row : dayRows) {
            sum += rowActualQtyM(headers, row);
        }
        return sum;
    }

    static double rowActualQtyM(List<String> headers, List<String> row) {
        int iAct = colIdx(headers, COL_ACTUAL_QTY);
        if (iAct >= 0) {
            double v = parseDouble(cellAt(row, iAct));
            if (Math.abs(v) > 1e-12) {
                return v;
            }
        }
        int iCum = colIdx(headers, COL_CUM_DONE);
        if (iCum >= 0) {
            double v = parseDouble(cellAt(row, iCum));
            if (Math.abs(v) > 1e-12) {
                return v;
            }
        }
        int iConv = colIdx(headers, COL_QTY_CONV);
        if (iConv >= 0) {
            double conv = parseDouble(cellAt(row, iConv));
            if (conv > 1e-12) {
                return conv * parseCompletionPct(headers, row) / 100.0;
            }
        }
        return 0.0;
    }

    static double sumAladdinPlanQtyM(
            List<String> headers,
            List<List<String>> rows,
            String machine,
            String planDateKey) {
        int mkIdx = colIdx(headers, COL_MACHINE);
        if (mkIdx < 0 || planDateKey == null || planDateKey.isBlank()) {
            return 0.0;
        }
        Integer dateCol = null;
        for (Map.Entry<Integer, String> e : dateColumnIndices(headers).entrySet()) {
            if (planDateKey.equals(normalizeDateHeader(e.getValue()))) {
                dateCol = e.getKey();
                break;
            }
        }
        if (dateCol == null) {
            return 0.0;
        }
        String mkNorm = normalizeEquipmentKey(machine);
        double sum = 0.0;
        for (List<String> row : rows) {
            if (!mkNorm.equals(normalizeEquipmentKey(cellAt(row, mkIdx)))) {
                continue;
            }
            sum += parseDouble(cellAt(row, dateCol));
        }
        return sum;
    }

    /** 完了率（%）= 当日実績(m) / 当日アラジン計画(m) × 100。計画0のときは0%。 */
    static double completionPctFromActualAndPlan(double actualQtyM, double aladdinPlanM) {
        if (aladdinPlanM <= 1e-12) {
            return 0.0;
        }
        return clampPct(actualQtyM / aladdinPlanM * 100.0);
    }

    private static Optional<EquipmentMachineStatus.ActualTaskRow> applyDayCompletionPct(
            Optional<EquipmentMachineStatus.ActualTaskRow> task,
            List<String> actHeaders,
            List<List<String>> dayRows,
            List<String> alHeaders,
            List<List<String>> alRows,
            String machine,
            String actualDateKey) {
        if (task.isEmpty()) {
            return task;
        }
        double actualQty = sumActualQtyM(actHeaders, dayRows);
        double plan = sumAladdinPlanQtyM(alHeaders, alRows, machine, actualDateKey);
        double pct = completionPctFromActualAndPlan(actualQty, plan);
        EquipmentMachineStatus.ActualTaskRow t = task.get();
        return Optional.of(
                new EquipmentMachineStatus.ActualTaskRow(
                        t.requestNo(),
                        t.processName(),
                        t.qtyConvM(),
                        pct,
                        t.memberRaw(),
                        t.startDateTime(),
                        t.endDateTime()));
    }

    private static EquipmentMachineStatus.ActualTaskRow enrichActualTaskMember(
            EquipmentMachineStatus.ActualTaskRow task,
            List<String> alHeaders,
            List<List<String>> alRows,
            List<String> disHeaders,
            List<List<String>> disRows,
            String machine,
            String actualDateKey,
            String planDateKey) {
        if (task.memberRaw() != null
                && !task.memberRaw().isBlank()
                && PersonNameHeuristics.looksLikePersonName(task.memberRaw())) {
            return task;
        }
        String member = lookupMemberFromAladdin(alHeaders, alRows, machine, task, actualDateKey);
        if (member.isBlank() && !actualDateKey.equals(planDateKey)) {
            member = lookupMemberFromAladdin(alHeaders, alRows, machine, task, planDateKey);
        }
        if (member.isBlank()) {
            member = lookupDispatchMember(disHeaders, disRows, machine, task, actualDateKey);
        }
        if (member.isBlank() && !actualDateKey.equals(planDateKey)) {
            member = lookupDispatchMember(disHeaders, disRows, machine, task, planDateKey);
        }
        if (member.isBlank()) {
            member = lookupDispatchMember(disHeaders, disRows, machine, task, null);
        }
        if (member.isBlank()) {
            return task;
        }
        return new EquipmentMachineStatus.ActualTaskRow(
                task.requestNo(),
                task.processName(),
                task.qtyConvM(),
                task.completionPct(),
                member,
                task.startDateTime(),
                task.endDateTime());
    }

    private static String lookupMemberFromAladdin(
            List<String> alHeaders,
            List<List<String>> alRows,
            String machine,
            EquipmentMachineStatus.ActualTaskRow task,
            String dateKey) {
        return AladdinShapedPlanMemberLookup.lookup(
                alHeaders,
                alRows,
                machine,
                task.requestNo(),
                task.processName(),
                dateKey);
    }

    private static String lookupDispatchMember(
            List<String> headers,
            List<List<String>> rows,
            String machine,
            EquipmentMachineStatus.ActualTaskRow task,
            String dateKey) {
        int mkIdx = colIdx(headers, ResultDispatchSchema.COL_MACHINE);
        int tidIdx = colIdx(headers, COL_REQUEST);
        int procIdx = colIdx(headers, ResultDispatchSchema.COL_PROCESS);
        int dateIdx = colIdx(headers, ResultDispatchSchema.COL_DISPATCH_DATE);
        int memberIdx = colIdx(headers, COL_MEMBER);
        if (mkIdx < 0 || tidIdx < 0 || memberIdx < 0) {
            return "";
        }
        String mkNorm = normalizeEquipmentKey(machine);
        String tid = task.requestNo() != null ? task.requestNo().strip() : "";
        String procKey = AladdinShapedPlanQtyLookup.normalizeProcessNameForRuleMatch(task.processName());
        for (List<String> row : rows) {
            if (!mkNorm.equals(normalizeEquipmentKey(cellAt(row, mkIdx)))) {
                continue;
            }
            if (!tid.equals(cellAt(row, tidIdx).strip())) {
                continue;
            }
            if (dateIdx >= 0 && dateKey != null && !dateKey.isBlank()) {
                if (!dateKey.equals(normalizeDateHeader(cellAt(row, dateIdx)))) {
                    continue;
                }
            }
            String rowProc =
                    procIdx >= 0
                            ? AladdinShapedPlanQtyLookup.normalizeProcessNameForRuleMatch(
                                    cellAt(row, procIdx))
                            : "";
            if (!procKey.isEmpty()
                    && !rowProc.isEmpty()
                    && !procKey.equals(rowProc)) {
                continue;
            }
            String member = cellAt(row, memberIdx).strip();
            if (PersonNameHeuristics.looksLikePersonName(member)) {
                return member;
            }
        }
        return "";
    }

    private static Optional<EquipmentMachineStatus.ActualTaskRow> pickLatestActualTask(
            List<String> headers, List<List<String>> dayRows) {
        if (dayRows == null || dayRows.isEmpty()) {
            return Optional.empty();
        }
        List<String> best = null;
        LocalDateTime bestDt = null;
        for (List<String> row : dayRows) {
            LocalDateTime dt = parseDateTime(cellAt(row, colIdx(headers, COL_START_DT)));
            if (dt == null) {
                LocalDate d = rowActualDate(headers, row);
                if (d != null) {
                    dt = d.atStartOfDay();
                }
            }
            if (best == null || compareDateTime(dt, bestDt) > 0) {
                best = row;
                bestDt = dt;
            }
        }
        if (best == null) {
            return Optional.empty();
        }
        int iReq = colIdx(headers, COL_REQUEST);
        int iProc = colIdx(headers, COL_PROCESS);
        int iConv = colIdx(headers, COL_QTY_CONV);
        int iStart = colIdx(headers, COL_START_DT);
        int iEnd = colIdx(headers, COL_END_DT);
        double pct = parseCompletionPct(headers, best);
        return Optional.of(
                new EquipmentMachineStatus.ActualTaskRow(
                        cellAt(best, iReq).strip(),
                        cellAt(best, iProc).strip(),
                        parseDouble(cellAt(best, iConv)),
                        pct,
                        readMemberFromRow(headers, best),
                        cellAt(best, iStart).strip(),
                        cellAt(best, iEnd).strip()));
    }

    private static String readMemberFromRow(List<String> headers, List<String> row) {
        for (String col : ACTUAL_MEMBER_COLUMNS) {
            int idx = colIdx(headers, col);
            if (idx >= 0) {
                String v = cellAt(row, idx).strip();
                if (PersonNameHeuristics.looksLikePersonName(v)) {
                    return v;
                }
            }
        }
        return "";
    }

    private static Map<String, List<List<String>>> groupActualRowsByMachine(
            List<String> headers, List<List<String>> rows, String dateKey) {
        int iMach = colIdx(headers, COL_MACHINE);
        if (iMach < 0 || dateKey == null) {
            return Map.of();
        }
        Map<String, List<List<String>>> byMachine = new LinkedHashMap<>();
        for (List<String> row : rows) {
            String mk = displayMachineKey(cellAt(row, iMach));
            if (mk.isEmpty()) {
                continue;
            }
            LocalDate d = rowActualDate(headers, row);
            if (d == null || !formatDateKey(d).equals(dateKey)) {
                continue;
            }
            byMachine.computeIfAbsent(mk, k -> new ArrayList<>()).add(row);
        }
        return byMachine;
    }

    private static Set<String> allMachinesInTable(
            List<String> headers, List<List<String>> rows, String machineColumn) {
        int mkIdx = colIdx(headers, machineColumn);
        if (mkIdx < 0 || rows == null || rows.isEmpty()) {
            return Set.of();
        }
        Set<String> out = new LinkedHashSet<>();
        for (List<String> row : rows) {
            String mk = displayMachineKey(cellAt(row, mkIdx));
            if (!mk.isEmpty()) {
                out.add(mk);
            }
        }
        return out;
    }

    private static List<EquipmentMachineStatus.PlanLine> collectAladdinPlans(
            List<String> headers,
            List<List<String>> rows,
            String machine,
            String planDateKey) {
        int mkIdx = colIdx(headers, COL_MACHINE);
        int tidIdx = colIdx(headers, COL_REQUEST);
        int procIdx = colIdx(headers, COL_PROCESS);
        if (mkIdx < 0 || tidIdx < 0 || planDateKey == null) {
            return List.of();
        }
        Integer dateCol = null;
        for (Map.Entry<Integer, String> e : dateColumnIndices(headers).entrySet()) {
            if (planDateKey.equals(normalizeDateHeader(e.getValue()))) {
                dateCol = e.getKey();
                break;
            }
        }
        if (dateCol == null) {
            return List.of();
        }
        String mkNorm = normalizeEquipmentKey(machine);
        List<EquipmentMachineStatus.PlanLine> out = new ArrayList<>();
        for (List<String> row : rows) {
            if (!mkNorm.equals(normalizeEquipmentKey(cellAt(row, mkIdx)))) {
                continue;
            }
            double qty = parseDouble(cellAt(row, dateCol));
            if (Math.abs(qty) <= 1e-12) {
                continue;
            }
            out.add(
                    new EquipmentMachineStatus.PlanLine(
                            cellAt(row, tidIdx).strip(),
                            procIdx >= 0 ? cellAt(row, procIdx).strip() : "",
                            ResultDispatchNormalizer.formatQty(qty)));
        }
        out.sort(planLineComparator());
        return List.copyOf(out);
    }

    private static List<EquipmentMachineStatus.PlanLine> collectDispatchPlans(
            List<String> headers,
            List<List<String>> rows,
            String machine,
            String planDateKey) {
        int mkIdx = colIdx(headers, ResultDispatchSchema.COL_MACHINE);
        int tidIdx = colIdx(headers, COL_REQUEST);
        int procIdx = colIdx(headers, ResultDispatchSchema.COL_PROCESS);
        int dateIdx = colIdx(headers, ResultDispatchSchema.COL_DISPATCH_DATE);
        int qtyIdx = colIdx(headers, ResultDispatchSchema.COL_DISPATCH_QTY);
        if (mkIdx < 0 || tidIdx < 0 || dateIdx < 0 || qtyIdx < 0 || planDateKey == null) {
            return List.of();
        }
        String mkNorm = normalizeEquipmentKey(machine);
        List<EquipmentMachineStatus.PlanLine> out = new ArrayList<>();
        for (List<String> row : rows) {
            if (!mkNorm.equals(normalizeEquipmentKey(cellAt(row, mkIdx)))) {
                continue;
            }
            if (!planDateKey.equals(normalizeDateHeader(cellAt(row, dateIdx)))) {
                continue;
            }
            double qty = parseDouble(cellAt(row, qtyIdx));
            if (Math.abs(qty) <= 1e-12) {
                continue;
            }
            out.add(
                    new EquipmentMachineStatus.PlanLine(
                            cellAt(row, tidIdx).strip(),
                            procIdx >= 0 ? cellAt(row, procIdx).strip() : "",
                            ResultDispatchNormalizer.formatQty(qty)));
        }
        out.sort(planLineComparator());
        return List.copyOf(out);
    }

    private static Comparator<EquipmentMachineStatus.PlanLine> planLineComparator() {
        return Comparator.comparing(EquipmentMachineStatus.PlanLine::requestNo, MACHINE_COLLATOR)
                .thenComparing(EquipmentMachineStatus.PlanLine::processName, MACHINE_COLLATOR);
    }

    private static Map<Integer, String> dateColumnIndices(List<String> headers) {
        Map<Integer, String> out = new LinkedHashMap<>();
        Pattern p = Pattern.compile("\\d{4}/\\d{2}/\\d{2}");
        for (int i = 0; i < headers.size(); i++) {
            String h = headers.get(i);
            if (h != null && p.matcher(h).matches()) {
                out.put(i, h);
            }
        }
        return out;
    }

    private static String displayMachineKey(String raw) {
        String n = normalizeEquipmentKey(raw);
        return n.isEmpty() ? "" : raw.strip();
    }

    private static String normalizeEquipmentKey(String val) {
        if (val == null || val.isBlank()) {
            return "";
        }
        String t = java.text.Normalizer.normalize(val, java.text.Normalizer.Form.NFKC);
        t = t.replace('\u00a0', ' ').replace('\u3000', ' ');
        t = t.replaceAll("[\u200b\u200c\u200d\ufeff]", "");
        return t.replaceAll("\\s+", " ").strip();
    }

    private static String normalizeDateHeader(String raw) {
        if (raw == null || raw.isBlank()) {
            return "";
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
                return String.format(Locale.ROOT, "%04d/%02d/%02d", y, mo, d);
            }
        } catch (NumberFormatException ignored) {
            // fall through
        }
        return "";
    }

    private static String formatDateKey(LocalDate d) {
        return d != null ? d.format(DateTimeFormatter.ofPattern("yyyy/MM/dd")) : "";
    }

    private static LocalDate parseDateKey(String raw) {
        String n = normalizeDateHeader(raw);
        if (n.isEmpty()) {
            return null;
        }
        try {
            return LocalDate.parse(n, DateTimeFormatter.ofPattern("yyyy/MM/dd"));
        } catch (DateTimeParseException e) {
            return null;
        }
    }

    private static LocalDate parseDatePrefix(String raw) {
        if (raw == null || raw.isBlank()) {
            return null;
        }
        String s = raw.strip();
        if (s.length() >= 10) {
            return parseDateKey(s.substring(0, 10));
        }
        return parseDateKey(s);
    }

    private static LocalDateTime parseDateTime(String raw) {
        if (raw == null || raw.isBlank()) {
            return null;
        }
        String s = raw.strip();
        try {
            if (s.length() >= 16) {
                return LocalDateTime.parse(s.substring(0, 16), DT_OUT);
            }
            LocalDate d = parseDateKey(s);
            return d != null ? d.atStartOfDay() : null;
        } catch (DateTimeParseException e) {
            return null;
        }
    }

    private static int compareDateTime(LocalDateTime a, LocalDateTime b) {
        if (a == null && b == null) {
            return 0;
        }
        if (a == null) {
            return -1;
        }
        if (b == null) {
            return 1;
        }
        return a.compareTo(b);
    }

    private static double parsePctCell(String raw) {
        if (raw == null || raw.isBlank()) {
            return Double.NaN;
        }
        String s = raw.strip();
        Matcher m = PCT_SUFFIX.matcher(s);
        if (m.find()) {
            s = s.substring(0, m.start()).strip();
            double v = parseDouble(s);
            return Double.isFinite(v) ? v : Double.NaN;
        }
        double v = parseDouble(s);
        if (!Double.isFinite(v)) {
            return Double.NaN;
        }
        if (v >= 0.0 && v <= 1.0) {
            return v * 100.0;
        }
        return v;
    }

    private static double clampPct(double v) {
        if (!Double.isFinite(v)) {
            return 0.0;
        }
        return Math.max(0.0, Math.min(100.0, v));
    }

    private static double parseDouble(String s) {
        if (s == null || s.isBlank()) {
            return 0.0;
        }
        try {
            return Double.parseDouble(s.strip().replace(",", ""));
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }

    private static int colIdx(List<String> headers, String title) {
        if (headers == null || title == null) {
            return -1;
        }
        for (int i = 0; i < headers.size(); i++) {
            if (title.equals(headers.get(i))) {
                return i;
            }
        }
        return -1;
    }

    private static String cellAt(List<String> row, int idx) {
        return (idx >= 0 && row != null && idx < row.size() && row.get(idx) != null)
                ? row.get(idx)
                : "";
    }

    private static List<String> copyList(List<String> in) {
        return in != null ? List.copyOf(in) : List.of();
    }

    private static List<List<String>> copyMatrix(List<List<String>> in) {
        if (in == null) {
            return List.of();
        }
        List<List<String>> out = new ArrayList<>(in.size());
        for (List<String> row : in) {
            out.add(row != null ? List.copyOf(row) : List.of());
        }
        return List.copyOf(out);
    }
}
