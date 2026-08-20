package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import jp.co.pm.ai.desktop.io.MasterTeamCombinationTableReader;

/** MASTER 4 シートの編集ロック・検証・列幅・組み合わせ行色。格子 UI から独立して単体試験する。 */
public final class MasterDispatchSheetEditRules {

    public enum SheetKind {
        SKILLS,
        NEED,
        SPEED,
        COMBINATIONS
    }

    static final int EXTRA_ROWS = 20;
    static final int EXTRA_COLS = 4;

    static final double COL_WIDTH_MIN = 110.0;
    static final double COL_WIDTH_MAX = 280.0;
    static final double COL_WIDTH_EMPTY = 56.0;

    private static final Pattern OP_AS = Pattern.compile("^(OP|AS)(\\d*)$", Pattern.CASE_INSENSITIVE);

    private static final String[] COMBO_BG = {
        "#d4edd4", "#cfe8f3", "#fde4cc", "#e6d4f0", "#f8d7da",
        "#fff2cc", "#d6eadf", "#dce3f0", "#f5e6c8", "#e8d5c4"
    };

    private MasterDispatchSheetEditRules() {}

    public static List<List<String>> skipFilterRow(List<List<String>> gridIncludingFilter) {
        if (gridIncludingFilter == null || gridIncludingFilter.isEmpty()) {
            return List.of();
        }
        return List.copyOf(gridIncludingFilter.subList(1, gridIncludingFilter.size()));
    }

    public static List<Double> preferredColumnWidths(List<List<String>> dataRows, int colCount) {
        int cols = Math.max(1, colCount);
        double[] max = new double[cols];
        if (dataRows != null) {
            for (List<String> row : dataRows) {
                if (row == null) {
                    continue;
                }
                for (int c = 0; c < cols && c < row.size(); c++) {
                    max[c] = Math.max(max[c], measureTextWidth(row.get(c)));
                }
            }
        }
        List<Double> out = new ArrayList<>(cols);
        for (int c = 0; c < cols; c++) {
            if (max[c] <= 28.0) {
                out.add(COL_WIDTH_EMPTY);
            } else {
                out.add(Math.min(COL_WIDTH_MAX, Math.max(COL_WIDTH_MIN, max[c])));
            }
        }
        return List.copyOf(out);
    }

    public static String comboRowStyle(String process, String machine) {
        return comboRowStyle(process, machine, "");
    }

    public static String comboRowStyle(String process, String machine, String comboCell) {
        String key = MasterTeamCombinationTableReader.comboKeyFromCells(process, machine, comboCell);
        if (key.isEmpty()) {
            return "";
        }
        int idx = Math.floorMod(key.hashCode(), COMBO_BG.length);
        String bg = COMBO_BG[idx];
        return "-fx-background-color: "
                + bg
                + "; -fx-control-inner-background: "
                + bg
                + "; -fx-text-fill: #111111;";
    }

    public static boolean isEditable(SheetKind kind, int dataRow, int col, List<List<String>> rows) {
        if (kind == null || dataRow < 0 || col < 0) {
            return false;
        }
        List<List<String>> src = rows != null ? rows : List.of();
        return switch (kind) {
            case SKILLS -> isSkillsEditable(dataRow, col);
            case NEED -> isNeedEditable(dataRow, col, src);
            case SPEED -> isSpeedEditable(dataRow, col, src);
            case COMBINATIONS -> isCombinationsEditable(dataRow, col, src);
        };
    }

    public static List<String> validateForSave(SheetKind kind, List<List<String>> rows) {
        List<List<String>> src = rows != null ? rows : List.of();
        return switch (kind) {
            case SKILLS -> validateSkills(src);
            case NEED -> validateNeed(src);
            case SPEED -> validateSpeed(src);
            case COMBINATIONS -> validateCombinations(src);
        };
    }

    public static List<List<String>> normalizeOnExtract(SheetKind kind, List<List<String>> rows) {
        List<List<String>> src = rows != null ? rows : List.of();
        return switch (kind) {
            case SKILLS -> normalizeSkills(src);
            case COMBINATIONS -> normalizeCombinations(src);
            default -> copyRows(src);
        };
    }

    public static boolean isInvalidValue(SheetKind kind, int dataRow, int col, List<List<String>> rows) {
        if (kind == null || dataRow < 0 || col < 0 || rows == null) {
            return false;
        }
        String v = cell(rows, dataRow, col);
        if (v.isEmpty()) {
            return false;
        }
        return switch (kind) {
            case SKILLS -> dataRow >= 2 && col >= 1 && parseOpAs(v) == null;
            case NEED -> isNeedValueInvalid(dataRow, col, rows, v);
            case SPEED -> isSpeedNumericCell(dataRow, col, rows) && !isPositiveNumber(v);
            case COMBINATIONS -> isCombinationsNumericInvalid(dataRow, col, rows, v);
        };
    }

    private static boolean isNeedValueInvalid(int dataRow, int col, List<List<String>> rows, String v) {
        if (!isNeedHeadcountCell(dataRow, col, rows)) {
            return false;
        }
        String a = cell(rows, dataRow, 0);
        if (a.startsWith("特別指定")) {
            return !isIntegerInRange(v, 1, 99);
        }
        if (a.contains("追加人数") || a.contains("余剰") || isNeedSurplusRow(dataRow, rows)) {
            return !isIntegerInRange(v, 0, 50);
        }
        return !isNonNegativeInteger(v);
    }

    public static int headerIndex(List<String> header, String... names) {
        if (header == null || names == null) {
            return -1;
        }
        for (int i = 0; i < header.size(); i++) {
            String h = header.get(i) != null ? header.get(i).strip().replace("組合せ", "組み合わせ") : "";
            for (String n : names) {
                if (n != null && h.equals(n)) {
                    return i;
                }
            }
        }
        return -1;
    }

    static String cell(List<List<String>> rows, int r, int c) {
        if (rows == null || r < 0 || r >= rows.size()) {
            return "";
        }
        List<String> row = rows.get(r);
        if (row == null || c < 0 || c >= row.size() || row.get(c) == null) {
            return "";
        }
        return row.get(c).strip();
    }

    private static boolean isSkillsEditable(int dataRow, int col) {
        if (dataRow <= 1 && col == 0) {
            return false;
        }
        return true;
    }

    private static boolean isNeedEditable(int dataRow, int col, List<List<String>> rows) {
        if (col == 0 && isNeedStructureLabel(cell(rows, dataRow, 0))) {
            return false;
        }
        return true;
    }

    private static boolean isSpeedEditable(int dataRow, int col, List<List<String>> rows) {
        if (col == 0 && isSpeedStructureLabel(cell(rows, dataRow, 0))) {
            return false;
        }
        return true;
    }

    private static boolean isCombinationsEditable(int dataRow, int col, List<List<String>> rows) {
        if (dataRow == 0) {
            return false;
        }
        int comboCol = headerIndex(headerRow(rows), "工程+機械", "工程＋機械");
        return comboCol < 0 || col != comboCol;
    }

    private static boolean isNeedStructureLabel(String a) {
        if (a.isEmpty()) {
            return false;
        }
        return a.equals("工程名")
                || a.equals("機械名")
                || a.contains("必須人数")
                || a.contains("必要人数")
                || a.contains("追加人数")
                || a.contains("余剰")
                || a.startsWith("特別指定");
    }

    private static boolean isSpeedStructureLabel(String a) {
        if (a.isEmpty()) {
            return false;
        }
        return a.equals("工程名")
                || a.equals("機械名")
                || a.contains("基本速度")
                || a.contains("実稼働比率")
                || a.startsWith("特別指定");
    }

    private static List<String> validateSkills(List<List<String>> rows) {
        List<String> errors = new ArrayList<>();
        int width = widthOf(rows);
        for (int r = 2; r < rows.size(); r++) {
            for (int c = 1; c < width; c++) {
                String v = cell(rows, r, c);
                if (!v.isEmpty() && parseOpAs(v) == null) {
                    errors.add("skills 交差セルは空か OP/AS+優先度のみです（例 OP1）。行"
                            + (r + 1)
                            + " 列"
                            + (c + 1)
                            + ": "
                            + v);
                }
            }
        }
        for (int c = 1; c < width; c++) {
            Map<Integer, String> seen = new HashMap<>();
            for (int r = 2; r < rows.size(); r++) {
                OpAs parsed = parseOpAs(cell(rows, r, c));
                if (parsed == null) {
                    continue;
                }
                String prev = seen.put(parsed.prio, cell(rows, r, 0));
                if (prev != null) {
                    errors.add("skills 同一列の優先度が重複しています（列"
                            + (c + 1)
                            + " 優先度"
                            + parsed.prio
                            + "）。");
                    break;
                }
            }
        }
        return List.copyOf(errors);
    }

    private static List<String> validateNeed(List<List<String>> rows) {
        List<String> errors = new ArrayList<>();
        int width = widthOf(rows);
        for (int r = 0; r < rows.size(); r++) {
            for (int c = 3; c < width; c++) {
                if (!isNeedHeadcountCell(r, c, rows)) {
                    continue;
                }
                String v = cell(rows, r, c);
                if (v.isEmpty()) {
                    continue;
                }
                String a = cell(rows, r, 0);
                if (a.startsWith("特別指定") && !isIntegerInRange(v, 1, 99)) {
                    errors.add("need 特別指定の人数は 1〜99 です。行" + (r + 1) + ": " + v);
                } else if ((a.contains("追加人数") || a.contains("余剰") || isNeedSurplusRow(r, rows))
                        && !isIntegerInRange(v, 0, 50)) {
                    errors.add("need 配台時追加人数は 0〜50 です。行" + (r + 1) + ": " + v);
                } else if (!isNonNegativeInteger(v)) {
                    errors.add("need の人数は 0 以上の整数です。行" + (r + 1) + ": " + v);
                }
            }
        }
        return List.copyOf(errors);
    }

    private static List<String> validateSpeed(List<List<String>> rows) {
        List<String> errors = new ArrayList<>();
        int width = widthOf(rows);
        for (int r = 0; r < rows.size(); r++) {
            for (int c = 3; c < width; c++) {
                if (!isSpeedNumericCell(r, c, rows)) {
                    continue;
                }
                String v = cell(rows, r, c);
                if (!v.isEmpty() && !isPositiveNumber(v)) {
                    errors.add("speed の基本速度・実稼働比率は数値です。行" + (r + 1) + ": " + v);
                }
            }
        }
        return List.copyOf(errors);
    }

    private static List<String> validateCombinations(List<List<String>> rows) {
        List<String> errors = new ArrayList<>();
        if (rows.isEmpty()) {
            return List.of();
        }
        List<String> header = headerRow(rows);
        int prioCol = headerIndex(header, "組み合わせ優先度", "組合せ優先度");
        int reqCol = headerIndex(header, "必須人数", "必要人数");
        for (int r = 1; r < rows.size(); r++) {
            if (prioCol >= 0) {
                String v = cell(rows, r, prioCol);
                if (!v.isEmpty() && !isPositiveInteger(v)) {
                    errors.add("組み合わせ優先度は 1 以上の整数です。行" + (r + 1) + ": " + v);
                }
            }
            if (reqCol >= 0) {
                String v = cell(rows, r, reqCol);
                if (!v.isEmpty() && !isNonNegativeInteger(v)) {
                    errors.add("必須人数は 0 以上の整数です。行" + (r + 1) + ": " + v);
                }
            }
        }
        return List.copyOf(errors);
    }

    private static boolean isNeedHeadcountCell(int dataRow, int col, List<List<String>> rows) {
        if (col < 3) {
            return false;
        }
        String a = cell(rows, dataRow, 0);
        if (a.contains("必須人数") || a.contains("必要人数") || a.contains("追加人数") || a.contains("余剰")) {
            return true;
        }
        return isNeedSurplusRow(dataRow, rows);
    }

    private static boolean isNeedSurplusRow(int dataRow, List<List<String>> rows) {
        int base = findFirstRowContaining(rows, "必須人数", "必要人数");
        return base >= 0 && dataRow == base + 1;
    }

    private static boolean isSpeedNumericCell(int dataRow, int col, List<List<String>> rows) {
        if (col < 3) {
            return false;
        }
        String a = cell(rows, dataRow, 0);
        if (a.contains("基本速度") || a.contains("実稼働比率")) {
            return true;
        }
        return dataRow == 3 || dataRow == 4;
    }

    private static boolean isCombinationsNumericInvalid(
            int dataRow, int col, List<List<String>> rows, String v) {
        if (dataRow == 0) {
            return false;
        }
        List<String> header = headerRow(rows);
        int prioCol = headerIndex(header, "組み合わせ優先度", "組合せ優先度");
        int reqCol = headerIndex(header, "必須人数", "必要人数");
        if (col == prioCol) {
            return !isPositiveInteger(v);
        }
        if (col == reqCol) {
            return !isNonNegativeInteger(v);
        }
        return false;
    }

    private static List<List<String>> normalizeSkills(List<List<String>> rows) {
        List<List<String>> out = mutableCopy(rows);
        int width = widthOf(out);
        for (int r = 2; r < out.size(); r++) {
            List<String> row = out.get(r);
            while (row.size() < width) {
                row.add("");
            }
            for (int c = 1; c < width; c++) {
                String v = row.get(c) != null ? row.get(c) : "";
                OpAs parsed = parseOpAs(v);
                if (parsed != null) {
                    row.set(c, parsed.role + parsed.prio);
                }
            }
        }
        return freeze(out);
    }

    private static List<List<String>> normalizeCombinations(List<List<String>> rows) {
        List<List<String>> out = mutableCopy(rows);
        if (out.isEmpty()) {
            return List.of();
        }
        List<String> header = out.get(0);
        int procCol = headerIndex(header, "工程名");
        int machCol = headerIndex(header, "機械名");
        int comboCol = headerIndex(header, "工程+機械", "工程＋機械");
        if (procCol < 0 || machCol < 0 || comboCol < 0) {
            return freeze(out);
        }
        int width = Math.max(widthOf(out), comboCol + 1);
        for (int r = 1; r < out.size(); r++) {
            List<String> row = out.get(r);
            while (row.size() < width) {
                row.add("");
            }
            String proc = procCol < row.size() && row.get(procCol) != null ? row.get(procCol).strip() : "";
            String mach = machCol < row.size() && row.get(machCol) != null ? row.get(machCol).strip() : "";
            if (!proc.isEmpty() && !mach.isEmpty()) {
                row.set(comboCol, proc + "+" + mach);
            }
        }
        return freeze(out);
    }

    private static List<String> headerRow(List<List<String>> rows) {
        if (rows == null || rows.isEmpty() || rows.get(0) == null) {
            return List.of();
        }
        return rows.get(0);
    }

    private static int findFirstRowContaining(List<List<String>> rows, String... needles) {
        for (int r = 0; r < rows.size(); r++) {
            String a = cell(rows, r, 0);
            for (String n : needles) {
                if (n != null && a.contains(n)) {
                    return r;
                }
            }
        }
        return -1;
    }

    private static int widthOf(List<List<String>> rows) {
        int w = 0;
        for (List<String> row : rows) {
            if (row != null) {
                w = Math.max(w, row.size());
            }
        }
        return w;
    }

    static OpAs parseOpAs(String raw) {
        if (raw == null) {
            return null;
        }
        String t = raw.replaceAll("\\s+", "");
        if (t.isEmpty()) {
            return null;
        }
        Matcher m = OP_AS.matcher(t);
        if (!m.matches()) {
            return null;
        }
        String role = m.group(1).toUpperCase(Locale.ROOT);
        String digits = m.group(2);
        int prio = digits == null || digits.isEmpty() ? 1 : Integer.parseInt(digits);
        if (prio < 1) {
            prio = 1;
        }
        return new OpAs(role, prio);
    }

    private static boolean isIntegerInRange(String s, int min, int max) {
        String t = stripTrailingDotZero(s);
        try {
            int n = Integer.parseInt(t);
            return n >= min && n <= max;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    private static boolean isNonNegativeInteger(String s) {
        String t = stripTrailingDotZero(s);
        try {
            return Integer.parseInt(t) >= 0;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    private static boolean isPositiveInteger(String s) {
        String t = stripTrailingDotZero(s);
        try {
            return Integer.parseInt(t) >= 1;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    private static boolean isPositiveNumber(String s) {
        try {
            double d = Double.parseDouble(s.strip().replace(",", ""));
            return !Double.isNaN(d) && !Double.isInfinite(d);
        } catch (NumberFormatException e) {
            return false;
        }
    }

    private static String stripTrailingDotZero(String s) {
        String t = s.strip().replace(",", "");
        if (t.endsWith(".0")) {
            return t.substring(0, t.length() - 2);
        }
        return t;
    }

    private static double measureTextWidth(String raw) {
        if (raw == null || raw.isEmpty()) {
            return 0;
        }
        double w = 28.0;
        for (int i = 0; i < raw.length(); i++) {
            w += raw.charAt(i) > 127 ? 14.0 : 8.0;
        }
        return w;
    }

    private static List<List<String>> copyRows(List<List<String>> rows) {
        return freeze(mutableCopy(rows));
    }

    private static List<List<String>> mutableCopy(List<List<String>> rows) {
        List<List<String>> out = new ArrayList<>();
        for (List<String> row : rows) {
            List<String> copy = new ArrayList<>();
            if (row != null) {
                for (String c : row) {
                    copy.add(c != null ? c : "");
                }
            }
            out.add(copy);
        }
        return out;
    }

    private static List<List<String>> freeze(List<List<String>> rows) {
        List<List<String>> out = new ArrayList<>(rows.size());
        for (List<String> row : rows) {
            out.add(List.copyOf(row));
        }
        return List.copyOf(out);
    }

    record OpAs(String role, int prio) {}
}
