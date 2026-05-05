package jp.co.pm.ai.desktop.print;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.TreeSet;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import jp.co.pm.ai.desktop.debug.AgentDebugLog;
import jp.co.pm.ai.desktop.io.JsonTableIo.SheetTable;

/**
 * Builds {@link OperatorCardPage} from member_schedule sheets and 結果_配台表 rows.
 */
public final class OperatorCardDocumentBuilder {

    /** Cursor debug session: オペレーターカードメンバー欄 / Y5-1 追跡 */
    private static final String DEBUG_SESSION_OP_CARD = "72c80c";

    private static final String COL_TIME = "時間帯";
    private static final Set<String> NON_WORK_MARKERS =
            Set.of(
                    "休",
                    "休憩",
                    "勤務外",
                    "年休");

    private static final Pattern MD_COL =
            Pattern.compile("(\\d{1,2})/(\\d{1,2})");
    private static final Pattern SLOT_RANGE =
            Pattern.compile("(\\d{2}:\\d{2})-(\\d{2}:\\d{2})");

    /** 結果_配台表.json の配台日（例 {@code 2026/05/07}）および ISO（{@code 2026-05-07}）を解釈する。 */
    private static final DateTimeFormatter DISPATCH_DAY_SLASH =
            DateTimeFormatter.ofPattern("uuuu/MM/dd");

    private static final String REQ = "依頼NO";
    private static final String PROC = "工程名";
    private static final String MACH = "機械名";
    private static final String DAY = "配台日";
    private static final String QTY_DAY = "当日配台数量";
    private static final String QTY_CONV = "換算数量";

    private OperatorCardDocumentBuilder() {}

    /** ログ用: 修正前後の比較用（生セルを空白除去したハッシュ）。 */
    static String normalizeCellProbe(String cell) {
        if (cell == null) {
            return "";
        }
        return cell.trim().replaceAll("\\p{Zs}+", "");
    }

    /**
     * チーム照合用キー。セル原文の完全一致ではなく、依頼NO・工程・機械（{@link MemberScheduleWorkCellParser} と同じ解釈）に
     * 正規化してオペレーター間の表記ゆれを吸収する。
     */
    static String canonicalTeamCellKey(String rawCell) {
        if (rawCell == null) {
            return "";
        }
        String trimmed = rawCell.trim();
        if (trimmed.isEmpty()) {
            return "";
        }
        MemberScheduleWorkCellParser.ParsedWorkCell p =
                MemberScheduleWorkCellParser.parse(trimmed);
        String req = p.requestNo();
        if (req != null && !req.isEmpty()) {
            return normalizeKey(req)
                    + "|"
                    + normalizeKey(p.processName())
                    + "|"
                    + normalizeKey(p.machineName());
        }
        return normalizeKey(trimmed);
    }

    static boolean traceY51Slice(String cell) {
        return cell != null && cell.contains("Y5-1");
    }

    public static OperatorCardPage buildPage(
            String operatorName,
            Map<String, SheetTable> memberSheetsByOperator,
            List<Map<String, String>> dispatchRows,
            LocalDate startDate)
            throws OperatorCardBuildException {

        SheetTable opSheet = memberSheetsByOperator.get(operatorName);
        if (opSheet == null) {
            throw new OperatorCardBuildException("no sheet for operator: " + operatorName);
        }
        if (!opSheet.columns().contains(COL_TIME)) {
            throw new OperatorCardBuildException("sheet missing column " + COL_TIME);
        }

        List<String> threeCols = resolveThreeDayColumns(opSheet.columns(), startDate);
        if (threeCols.stream().anyMatch(Objects::isNull)) {
            throw new OperatorCardBuildException(
                    "could not resolve date columns for "
                            + startDate
                            + " .. "
                            + startDate.plusDays(2));
        }

        Map<String, Set<String>> teamMap =
                buildTeamMap(memberSheetsByOperator, threeCols);

        // #region agent log
        {
            int y51Keys = 0;
            for (String k : teamMap.keySet()) {
                if (k != null && k.contains("Y5-1")) {
                    y51Keys++;
                }
            }
            AgentDebugLog.appendStructured(
                    Map.of(),
                    DEBUG_SESSION_OP_CARD,
                    "H3",
                    "OperatorCardDocumentBuilder.buildPage",
                    "teamMap built",
                    Map.of(
                            "operatorSheets",
                            memberSheetsByOperator.size(),
                            "dateCols",
                            String.join("|", threeCols),
                            "teamMapSize",
                            teamMap.size(),
                            "y51KeyCount",
                            y51Keys));
        }
        // #endregion

        List<OperatorCardDaySection> days = new ArrayList<>();
        for (int i = 0; i < 3; i++) {
            LocalDate d = startDate.plusDays(i);
            String colKey = threeCols.get(i);
            List<OperatorCardTaskRow> rows =
                    buildDayRows(opSheet, dispatchRows, teamMap, colKey, d, operatorName);
            days.add(new OperatorCardDaySection(d, colKey, rows));
        }

        return new OperatorCardPage(operatorName, days);
    }

    /**
     * Finds header strings matching MM/dd for {@code startDate} and the following two days (same calendar year as
     * {@code startDate}, with rollover handled only when {@code startDate} month is December).
     */
    public static List<String> resolveThreeDayColumns(List<String> columns, LocalDate startDate) {
        List<String> out = new ArrayList<>(3);
        for (int i = 0; i < 3; i++) {
            LocalDate want = startDate.plusDays(i);
            String found = findColumnForDate(columns, want, startDate);
            out.add(found);
        }
        return out;
    }

    static String findColumnForDate(List<String> columns, LocalDate want, LocalDate anchor) {
        int y = anchor.getYear();
        for (String col : columns) {
            if (COL_TIME.equals(col)) {
                continue;
            }
            LocalDate p = parseColumnDate(col, y);
            if (want.equals(p)) {
                return col;
            }
            p = parseColumnDate(col, y + 1);
            if (want.equals(p)) {
                return col;
            }
        }
        return null;
    }

    static LocalDate parseColumnDate(String columnTitle, int year) {
        Matcher m = MD_COL.matcher(columnTitle != null ? columnTitle : "");
        if (!m.find()) {
            return null;
        }
        try {
            int mo = Integer.parseInt(m.group(1));
            int da = Integer.parseInt(m.group(2));
            return LocalDate.of(year, mo, da);
        } catch (DateTimeParseException | IllegalArgumentException ignored) {
            return null;
        }
    }

    static Map<String, Set<String>> buildTeamMap(
            Map<String, SheetTable> memberSheetsByOperator, List<String> dateCols) {
        Map<String, Set<String>> out = new LinkedHashMap<>();
        for (Map.Entry<String, SheetTable> en : memberSheetsByOperator.entrySet()) {
            String op = en.getKey();
            SheetTable st = en.getValue();
            if (!st.columns().contains(COL_TIME)) {
                continue;
            }
            for (Map<String, String> row : st.rows()) {
                String slot = row.get(COL_TIME);
                if (slot == null || slot.isBlank()) {
                    continue;
                }
                for (String dc : dateCols) {
                    String cell = row.getOrDefault(dc, "").trim();
                    if (cell.isEmpty() || isNonWork(cell)) {
                        continue;
                    }
                    String k = teamKey(dc, slot, cell);
                    out.computeIfAbsent(k, x -> new LinkedHashSet<>()).add(op);
                    // #region agent log
                    if (traceY51Slice(cell)) {
                        AgentDebugLog.appendStructured(
                                Map.of(),
                                DEBUG_SESSION_OP_CARD,
                                "H2",
                                "OperatorCardDocumentBuilder.buildTeamMap",
                                "teamMap put",
                                Map.of(
                                        "operator",
                                        op,
                                        "dateCol",
                                        dc,
                                        "slot",
                                        slot != null ? slot : "",
                                        "cellLen",
                                        cell.length(),
                                        "cellNormLen",
                                        normalizeCellProbe(cell).length(),
                                        "cellNormHash",
                                        Integer.toHexString(normalizeCellProbe(cell).hashCode()),
                                        "canonicalKeyHash",
                                        Integer.toHexString(
                                                canonicalTeamCellKey(cell).hashCode())));
                    }
                    // #endregion
                }
            }
        }
        return out;
    }

    static String teamKey(String dateCol, String timeSlot, String rawCell) {
        return dateCol + "\0" + timeSlot + "\0" + canonicalTeamCellKey(rawCell);
    }

    static boolean isNonWork(String cell) {
        String t = cell.trim();
        return t.isEmpty() || NON_WORK_MARKERS.contains(t);
    }

    static List<OperatorCardTaskRow> buildDayRows(
            SheetTable operatorSheet,
            List<Map<String, String>> dispatchRows,
            Map<String, Set<String>> teamMap,
            String dateCol,
            LocalDate date,
            String selfOperator) {

        List<Map<String, String>> lines = operatorSheet.rows();
        List<OperatorCardTaskRow> out = new ArrayList<>();
        int i = 0;
        while (i < lines.size()) {
            Map<String, String> row = lines.get(i);
            String slot = row.get(COL_TIME);
            String cell = row.getOrDefault(dateCol, "").trim();
            if (slot == null || slot.isBlank() || cell.isEmpty() || isNonWork(cell)) {
                i++;
                continue;
            }
            int j = i;
            while (j + 1 < lines.size()) {
                String nextCell = lines.get(j + 1).getOrDefault(dateCol, "").trim();
                if (!cell.equals(nextCell)) {
                    break;
                }
                j++;
            }
            String timeRange = mergeTimeRange(lines.get(i).get(COL_TIME), lines.get(j).get(COL_TIME));
            MemberScheduleWorkCellParser.ParsedWorkCell p =
                    MemberScheduleWorkCellParser.parse(cell);

            Set<String> members = new TreeSet<>();
            for (int k = i; k <= j; k++) {
                String sl = lines.get(k).get(COL_TIME);
                if (sl != null) {
                    String tk = teamKey(dateCol, sl, cell);
                    members.addAll(teamMap.getOrDefault(tk, Set.of()));
                }
            }
            if (members.isEmpty()) {
                members.add(selfOperator);
            }
            members.remove(selfOperator);
            String memberStr =
                    members.isEmpty() ? "—" : String.join("、", members);

            Map<String, String> dispatchHit = null;
            String qtyD = "";
            String qtyC = "";
            if (!p.requestNo().isEmpty()) {
                dispatchHit =
                        findDispatchRow(
                                dispatchRows,
                                date,
                                p.requestNo(),
                                p.processName(),
                                p.machineName());
                if (dispatchHit != null) {
                    qtyD = nz(dispatchHit.get(QTY_DAY));
                    qtyC = nz(dispatchHit.get(QTY_CONV));
                }
            }

            // #region agent log
            if (traceY51Slice(cell)) {
                List<Integer> perKeySizes = new ArrayList<>();
                List<String> slotSamples = new ArrayList<>();
                for (int k = i; k <= j; k++) {
                    String sl = lines.get(k).get(COL_TIME);
                    if (sl != null) {
                        String tk = teamKey(dateCol, sl, cell);
                        perKeySizes.add(teamMap.getOrDefault(tk, Set.of()).size());
                        slotSamples.add(sl);
                    }
                }
                AgentDebugLog.appendStructured(
                        Map.of(),
                        DEBUG_SESSION_OP_CARD,
                        "H1",
                        "OperatorCardDocumentBuilder.buildDayRows",
                        "Y5-1 row members resolved",
                        Map.of(
                                "selfOperator",
                                selfOperator,
                                "date",
                                date.toString(),
                                "slotRange",
                                String.join(
                                        " .. ",
                                        slotSamples.isEmpty()
                                                ? List.of("")
                                                : List.of(slotSamples.get(0), slotSamples.get(slotSamples.size() - 1))),
                                "perKeySizes",
                                perKeySizes.toString(),
                                "memberStr",
                                memberStr,
                                "memberStrLen",
                                memberStr.length(),
                                "cellNormHash",
                                Integer.toHexString(normalizeCellProbe(cell).hashCode()),
                                "canonicalKeyHash",
                                Integer.toHexString(canonicalTeamCellKey(cell).hashCode()),
                                "runId",
                                "post-fix"));
            }
            // #endregion

            // #region agent log
            if (traceY51Slice(cell) && !p.requestNo().isEmpty()) {
                AgentDebugLog.appendStructured(
                        Map.of(),
                        DEBUG_SESSION_OP_CARD,
                        "QTY",
                        "OperatorCardDocumentBuilder.buildDayRows",
                        "dispatch lookup for quantities",
                        Map.of(
                                "dateIso",
                                date.toString(),
                                "dispatchHit",
                                dispatchHit != null,
                                "qtyDLen",
                                qtyD.length(),
                                "qtyCLen",
                                qtyC.length()));
            }
            // #endregion

            out.add(
                    new OperatorCardTaskRow(
                            timeRange,
                            nz(p.processName()),
                            nz(p.machineName()),
                            nz(p.requestNo()),
                            qtyD.isEmpty() ? "—" : qtyD,
                            qtyC.isEmpty() ? "—" : qtyC,
                            memberStr));

            i = j + 1;
        }
        return out;
    }

    /**
     * 結果_配台表の「配台日」セル文字列を {@link LocalDate} にする。ISO とスラッシュ区切りの両方を許容する。
     */
    static LocalDate parseDispatchTableDay(String cell) {
        if (cell == null) {
            return null;
        }
        String t = cell.trim();
        if (t.isEmpty()) {
            return null;
        }
        try {
            return LocalDate.parse(t);
        } catch (DateTimeParseException ignored) {
        }
        try {
            return LocalDate.parse(t, DISPATCH_DAY_SLASH);
        } catch (DateTimeParseException ignored) {
        }
        try {
            return LocalDate.parse(t, DateTimeFormatter.ofPattern("uuuu/M/d"));
        } catch (DateTimeParseException ignored) {
        }
        return null;
    }

    static Map<String, String> findDispatchRow(
            List<Map<String, String>> dispatchRows,
            LocalDate date,
            String reqNo,
            String process,
            String machine) {

        for (Map<String, String> r : dispatchRows) {
            LocalDate rowDay = parseDispatchTableDay(r.get(DAY));
            if (rowDay == null || !date.equals(rowDay)) {
                continue;
            }
            if (!eqNorm(reqNo, r.get(REQ))) {
                continue;
            }
            if (!eqNorm(process, r.get(PROC))) {
                continue;
            }
            if (!eqNorm(machine, r.get(MACH))) {
                continue;
            }
            return r;
        }
        return null;
    }

    static boolean eqNorm(String a, String b) {
        return normalizeKey(a).equals(normalizeKey(b));
    }

    static String normalizeKey(String s) {
        if (s == null) {
            return "";
        }
        return s.trim().replaceAll("\\p{Zs}+", "");
    }

    static String nz(String s) {
        return s != null ? s : "";
    }

    static String mergeTimeRange(String firstSlot, String lastSlot) {
        Matcher mf = SLOT_RANGE.matcher(firstSlot != null ? firstSlot : "");
        Matcher ml = SLOT_RANGE.matcher(lastSlot != null ? lastSlot : "");
        if (!mf.find() || !ml.find()) {
            return firstSlot != null ? firstSlot : "";
        }
        return mf.group(1) + "-" + ml.group(2);
    }

    public static final class OperatorCardBuildException extends Exception {
        public OperatorCardBuildException(String message) {
            super(message);
        }
    }
}
