package jp.co.pm.ai.desktop.print;

import java.time.LocalDate;
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

import jp.co.pm.ai.desktop.io.JsonTableIo.SheetTable;

/**
 * Builds {@link OperatorCardPage} from member_schedule sheets and ????_?z??\ rows.
 */
public final class OperatorCardDocumentBuilder {

    private static final String COL_TIME = "\u6642\u9593\u5e2f";
    private static final Set<String> NON_WORK_MARKERS =
            Set.of(
                    "\u4f11",
                    "\u52e4\u52d9\u5916",
                    "\u5e74\u4f11");

    private static final Pattern MD_COL =
            Pattern.compile("(\\d{1,2})/(\\d{1,2})");
    private static final Pattern SLOT_RANGE =
            Pattern.compile("(\\d{2}:\\d{2})-(\\d{2}:\\d{2})");

    private static final String REQ = "\u4f9d\u983cNO";
    private static final String PROC = "\u5de5\u7a0b\u540d";
    private static final String MACH = "\u6a5f\u68b0\u540d";
    private static final String DAY = "\u914d\u53f0\u65e5";
    private static final String QTY_DAY = "\u5f53\u65e5\u914d\u53f0\u6570\u91cf";
    private static final String QTY_CONV = "\u63db\u7b97\u6570\u91cf";

    private OperatorCardDocumentBuilder() {}

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
                }
            }
        }
        return out;
    }

    static String teamKey(String dateCol, String timeSlot, String rawCell) {
        return dateCol + "\0" + timeSlot + "\0" + rawCell;
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
            String memberStr = String.join("\u3001", members);

            String qtyD = "";
            String qtyC = "";
            if (!p.requestNo().isEmpty()) {
                Map<String, String> hit =
                        findDispatchRow(dispatchRows, date, p.requestNo(), p.processName(), p.machineName());
                if (hit != null) {
                    qtyD = nz(hit.get(QTY_DAY));
                    qtyC = nz(hit.get(QTY_CONV));
                }
            }

            out.add(
                    new OperatorCardTaskRow(
                            timeRange,
                            nz(p.processName()),
                            nz(p.machineName()),
                            nz(p.requestNo()),
                            qtyD.isEmpty() ? "\u2014" : qtyD,
                            qtyC.isEmpty() ? "\u2014" : qtyC,
                            memberStr));

            i = j + 1;
        }
        return out;
    }

    static Map<String, String> findDispatchRow(
            List<Map<String, String>> dispatchRows,
            LocalDate date,
            String reqNo,
            String process,
            String machine) {

        String dayStr = date.toString();
        for (Map<String, String> r : dispatchRows) {
            if (!dayStr.equals(nz(r.get(DAY)))) {
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
        return s.trim().replaceAll("\\s+", "");
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
