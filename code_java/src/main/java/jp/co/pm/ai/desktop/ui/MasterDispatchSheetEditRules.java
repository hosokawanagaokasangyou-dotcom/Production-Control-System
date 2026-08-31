package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.BitSet;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
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

    /** 工程名・機械名など、データと区別する見出し行。 */
    public enum TitleRowKind {
        NONE,
        PROCESS,
        MACHINE,
        OTHER
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

    public static List<String> columnTitles(SheetKind kind, List<List<String>> rows, int colCount) {
        int cols = Math.max(1, colCount);
        List<String> titles = new ArrayList<>(cols);
        for (int c = 0; c < cols; c++) {
            titles.add("");
        }
        List<List<String>> src = rows != null ? rows : List.of();
        if (kind == SheetKind.COMBINATIONS) {
            List<String> header = headerRow(src);
            for (int c = 0; c < cols && c < header.size(); c++) {
                titles.set(c, header.get(c) != null ? header.get(c).strip() : "");
            }
            return List.copyOf(titles);
        }
        if (kind == SheetKind.SKILLS) {
            titles.set(0, "メンバー");
        } else {
            titles.set(0, "項目");
        }
        if (kind == SheetKind.NEED) {
            titles.set(1, "依頼NO条件");
            if (cols > 2) {
                titles.set(2, "備考");
            }
        }
        int procRow = findProcessHeaderRow(src);
        int machRow = findMachineHeaderRow(src);
        if (procRow < 0) {
            procRow = 0;
        }
        if (machRow < 0) {
            machRow = src.size() > 1 ? 1 : -1;
        }
        int firstEq = kind == SheetKind.NEED || kind == SheetKind.SPEED ? 3 : 1;
        for (int c = firstEq; c < cols; c++) {
            String proc = cell(src, procRow, c);
            String mach = machRow >= 0 ? cell(src, machRow, c) : "";
            titles.set(c, equipmentColumnTitle(proc, mach));
        }
        return List.copyOf(titles);
    }

    public static List<List<String>> displayRows(SheetKind kind, List<List<String>> rows) {
        List<List<String>> src = rows != null ? rows : List.of();
        List<List<String>> out = new ArrayList<>();
        for (int r = 0; r < src.size(); r++) {
            if (kind == SheetKind.COMBINATIONS && isColumnTitleSourceRow(kind, r, src)) {
                continue;
            }
            if (kind != SheetKind.COMBINATIONS && isNeedColumnCaptionRow(src, r)) {
                continue;
            }
            String a = cell(src, r, 0);
            if (isProcessOrMachineHeaderLabel(a) && !"工程名".equals(a) && !"機械名".equals(a)) {
                continue;
            }
            out.add(src.get(r));
        }
        return freeze(out);
    }

    /** skills / need / speed の工程名・機械名行を縦スクロール固定する件数（フィルタ行の次から）。 */
    public static int frozenTitleRowCount(SheetKind kind) {
        return kind == null || kind == SheetKind.COMBINATIONS ? 0 : 2;
    }

    public static boolean isSkillsSkillValueCell(int dataRow, int col, List<List<String>> rows) {
        return col >= 1 && !isColumnTitleSourceRow(SheetKind.SKILLS, dataRow, rows);
    }

    /**
     * 工程名+機械名の設備列を追加する。既にある組み合わせは何もしない。
     * need / speed は先頭 3 列（項目・依頼NO条件・備考）の右へ足す。
     */
    public static List<List<String>> addEquipmentColumn(
            SheetKind kind, List<List<String>> rows, String process, String machine) {
        String p = process != null ? process.strip() : "";
        String m = machine != null ? machine.strip() : "";
        if (p.isEmpty() || m.isEmpty() || kind == null || kind == SheetKind.COMBINATIONS) {
            return copyRows(rows != null ? rows : List.of());
        }
        List<List<String>> out = mutableCopy(rows != null ? rows : List.of());
        ensureProcessMachineHeaderRows(kind, out);
        int procRow = findProcessHeaderRow(out);
        int machRow = findMachineHeaderRow(out);
        if (procRow < 0 || machRow < 0) {
            return freeze(out);
        }
        int firstEq = kind == SheetKind.NEED || kind == SheetKind.SPEED ? 3 : 1;
        String want = MasterTeamCombinationTableReader.normalizedComboKey(p, m);
        int width = Math.max(widthOf(out), firstEq);
        padWidth(out, width);
        for (int c = firstEq; c < width; c++) {
            String have =
                    MasterTeamCombinationTableReader.normalizedComboKey(
                            cell(out, procRow, c), cell(out, machRow, c));
            if (!want.isEmpty() && want.equals(have)) {
                return freeze(out);
            }
        }
        int target = -1;
        for (int c = firstEq; c < width; c++) {
            if (cell(out, procRow, c).isEmpty() && cell(out, machRow, c).isEmpty()) {
                target = c;
                break;
            }
        }
        if (target < 0) {
            target = width;
            padWidth(out, target + 1);
        }
        setCell(out, procRow, target, p);
        setCell(out, machRow, target, m);
        return freeze(out);
    }

    public static boolean containsEquipmentColumn(
            SheetKind kind, List<List<String>> rows, String process, String machine) {
        String p = process != null ? process.strip() : "";
        String m = machine != null ? machine.strip() : "";
        if (p.isEmpty() || m.isEmpty() || kind == null || kind == SheetKind.COMBINATIONS) {
            return false;
        }
        List<List<String>> src = rows != null ? rows : List.of();
        int procRow = findProcessHeaderRow(src);
        int machRow = findMachineHeaderRow(src);
        if (procRow < 0 || machRow < 0) {
            return false;
        }
        int firstEq = kind == SheetKind.NEED || kind == SheetKind.SPEED ? 3 : 1;
        String want = MasterTeamCombinationTableReader.normalizedComboKey(p, m);
        if (want.isEmpty()) {
            return false;
        }
        int width = widthOf(src);
        for (int c = firstEq; c < width; c++) {
            String have =
                    MasterTeamCombinationTableReader.normalizedComboKey(
                            cell(src, procRow, c), cell(src, machRow, c));
            if (want.equals(have)) {
                return true;
            }
        }
        return false;
    }

    public static final String COL_EDIT_LOCK = "編集ロック";
    public static final String COL_ADDED_ROW = "追加行";
    public static final String LOCK_VALUE = "ロック";
    public static final String ADDED_FLAG = "1";
    public static final String ADDED_ROW_BG = "#f4c36a";

    private static final Pattern COMBO_MEMBER_ROLE_PREFIX =
            Pattern.compile("(?i)^(OP|AS)\\s*\\d*\\s+");

    public static List<List<String>> ensureCombinationMetaColumns(List<List<String>> rows) {
        List<List<String>> out = mutableCopy(rows != null ? rows : List.of());
        if (out.isEmpty()) {
            out.add(new ArrayList<>(List.of("組み合わせ行ID", "工程名", "機械名", "工程+機械")));
        }
        List<String> header = out.get(0);
        if (headerIndex(header, COL_EDIT_LOCK) < 0) {
            header.add(COL_EDIT_LOCK);
        }
        if (headerIndex(header, COL_ADDED_ROW) < 0) {
            header.add(COL_ADDED_ROW);
        }
        padWidth(out, header.size());
        return freeze(out);
    }

    public static boolean isCombinationRowLocked(List<List<String>> rows, int dataRow) {
        if (rows == null || dataRow <= 0 || dataRow >= rows.size()) {
            return false;
        }
        int lockCol = headerIndex(headerRow(rows), COL_EDIT_LOCK);
        return lockCol >= 0 && isLockFlag(cell(rows, dataRow, lockCol));
    }

    public static boolean isAddedCombinationRow(List<List<String>> rows, int dataRow) {
        if (rows == null || dataRow <= 0 || dataRow >= rows.size()) {
            return false;
        }
        int addedCol = headerIndex(headerRow(rows), COL_ADDED_ROW);
        return addedCol >= 0 && isAddedFlag(cell(rows, dataRow, addedCol));
    }

    public static boolean isCombinationLockColumn(List<String> header, int col) {
        return col >= 0 && headerIndex(header, COL_EDIT_LOCK) == col;
    }

    public static boolean isCombinationMemberColumn(List<String> header, int col) {
        if (header == null || col < 0 || col >= header.size() || header.get(col) == null) {
            return false;
        }
        String h = header.get(col).strip().replace("組合せ", "組み合わせ");
        return h.startsWith("メンバー");
    }

    public static List<List<String>> addCombinationRow(
            List<List<String>> rows, String process, String machine) {
        String p = process != null ? process.strip() : "";
        String m = machine != null ? machine.strip() : "";
        if (p.isEmpty() || m.isEmpty()) {
            return copyRows(rows != null ? rows : List.of());
        }
        String want = MasterTeamCombinationTableReader.normalizedComboKey(p, m);
        if (want.isEmpty()) {
            return copyRows(rows != null ? rows : List.of());
        }
        if (containsCombinationEquipment(rows, p, m)) {
            return copyRows(rows != null ? rows : List.of());
        }
        List<List<String>> out = mutableCopy(ensureCombinationMetaColumns(rows));
        List<String> header = out.get(0);
        int width = header.size();
        List<String> row = new ArrayList<>();
        while (row.size() < width) {
            row.add("");
        }
        int idCol = headerIndex(header, "組み合わせ行ID", "組合せ行ID", "インデックス");
        int procCol = headerIndex(header, "工程名");
        int machCol = headerIndex(header, "機械名");
        int comboCol = headerIndex(header, "工程+機械", "工程＋機械");
        int prioCol = headerIndex(header, "組み合わせ優先度", "組合せ優先度");
        int addedCol = headerIndex(header, COL_ADDED_ROW);
        if (idCol >= 0) {
            row.set(idCol, String.valueOf(nextCombinationRowId(out)));
        }
        if (procCol >= 0) {
            row.set(procCol, p);
        }
        if (machCol >= 0) {
            row.set(machCol, m);
        }
        if (comboCol >= 0) {
            row.set(comboCol, p + "+" + m);
        }
        if (prioCol >= 0) {
            row.set(prioCol, "1");
        }
        if (addedCol >= 0) {
            row.set(addedCol, ADDED_FLAG);
        }
        out.add(row);
        return freeze(out);
    }

    public static boolean containsCombinationEquipment(
            List<List<String>> rows, String process, String machine) {
        String want =
                MasterTeamCombinationTableReader.normalizedComboKey(
                        process != null ? process : "", machine != null ? machine : "");
        if (want.isEmpty() || rows == null || rows.size() < 2) {
            return false;
        }
        List<String> header = headerRow(rows);
        int procCol = headerIndex(header, "工程名");
        int machCol = headerIndex(header, "機械名");
        int comboCol = headerIndex(header, "工程+機械", "工程＋機械");
        for (int r = 1; r < rows.size(); r++) {
            if (isColumnTitleSourceRow(SheetKind.COMBINATIONS, r, rows)) {
                continue;
            }
            String proc = procCol >= 0 ? cell(rows, r, procCol) : "";
            String mach = machCol >= 0 ? cell(rows, r, machCol) : "";
            String combo = comboCol >= 0 ? cell(rows, r, comboCol) : "";
            String have = MasterTeamCombinationTableReader.comboKeyFromCells(proc, mach, combo);
            if (want.equals(have)) {
                return true;
            }
        }
        return false;
    }

    public static List<List<String>> deleteCombinationRows(
            List<List<String>> rows, Set<Integer> originalRowIndexes) {
        if (rows == null || rows.isEmpty() || originalRowIndexes == null || originalRowIndexes.isEmpty()) {
            return copyRows(rows != null ? rows : List.of());
        }
        List<List<String>> out = new ArrayList<>();
        for (int r = 0; r < rows.size(); r++) {
            if (r > 0
                    && originalRowIndexes.contains(r)
                    && !isCombinationRowLocked(rows, r)
                    && !isColumnTitleSourceRow(SheetKind.COMBINATIONS, r, rows)) {
                continue;
            }
            List<String> copy = new ArrayList<>();
            if (rows.get(r) != null) {
                for (String c : rows.get(r)) {
                    copy.add(c != null ? c : "");
                }
            }
            out.add(copy);
        }
        return freeze(out);
    }

    public static List<String> combinationMemberChoices(
            List<List<String>> skillsRows, String process, String machine, String current) {
        LinkedHashSet<String> items = new LinkedHashSet<>();
        items.add("");
        List<List<String>> skills = skillsRows != null ? skillsRows : List.of();
        int procRow = findProcessHeaderRow(skills);
        int machRow = findMachineHeaderRow(skills);
        String want = MasterTeamCombinationTableReader.normalizedComboKey(process, machine);
        int eqCol = -1;
        if (!want.isEmpty() && procRow >= 0 && machRow >= 0) {
            int width = widthOf(skills);
            for (int c = 1; c < width; c++) {
                String have =
                        MasterTeamCombinationTableReader.normalizedComboKey(
                                cell(skills, procRow, c), cell(skills, machRow, c));
                if (want.equals(have)) {
                    eqCol = c;
                    break;
                }
            }
        }
        if (eqCol >= 0) {
            for (int r = 0; r < skills.size(); r++) {
                if (isColumnTitleSourceRow(SheetKind.SKILLS, r, skills)) {
                    continue;
                }
                if (isProcessOrMachineHeaderLabel(cell(skills, r, 0))) {
                    continue;
                }
                OpAs parsed = parseOpAs(cell(skills, r, eqCol));
                if (parsed == null) {
                    continue;
                }
                String name = cell(skills, r, 0);
                if (name.isEmpty()) {
                    continue;
                }
                items.add(parsed.role() + " " + name);
            }
        }
        String cur = current != null ? current.strip() : "";
        if (!cur.isEmpty() && !items.contains(cur)) {
            items.add(cur);
        }
        return List.copyOf(items);
    }

    public static String combinationMemberName(String raw) {
        String t = raw != null ? raw.strip() : "";
        if (t.isEmpty()) {
            return "";
        }
        Matcher m = COMBO_MEMBER_ROLE_PREFIX.matcher(t);
        if (m.find()) {
            return m.replaceFirst("").strip();
        }
        return t;
    }

    public static List<String[]> skillsEquipmentPairs(List<List<String>> skillsRows) {
        List<String> titles =
                columnTitles(SheetKind.SKILLS, skillsRows, Math.max(1, widthOf(skillsRows != null ? skillsRows : List.of())));
        List<String[]> out = new ArrayList<>();
        LinkedHashSet<String> seen = new LinkedHashSet<>();
        for (int i = 1; i < titles.size(); i++) {
            String[] pm = splitEquipmentTitle(titles.get(i));
            if (pm[0].isEmpty() || pm[1].isEmpty()) {
                continue;
            }
            String key = MasterTeamCombinationTableReader.normalizedComboKey(pm[0], pm[1]);
            if (key.isEmpty() || !seen.add(key)) {
                continue;
            }
            out.add(new String[] {pm[0], pm[1]});
        }
        return List.copyOf(out);
    }

    private static int nextCombinationRowId(List<List<String>> rows) {
        int idCol = headerIndex(headerRow(rows), "組み合わせ行ID", "組合せ行ID", "インデックス");
        int max = 0;
        if (idCol < 0) {
            return 1;
        }
        for (int r = 1; r < rows.size(); r++) {
            String v = cell(rows, r, idCol);
            try {
                int n = Integer.parseInt(stripTrailingDotZero(v));
                if (n > max) {
                    max = n;
                }
            } catch (NumberFormatException ignored) {
                // skip
            }
        }
        return max + 1;
    }

    private static boolean isLockFlag(String v) {
        String t = v != null ? v.strip() : "";
        return LOCK_VALUE.equals(t) || "1".equals(t) || "true".equalsIgnoreCase(t);
    }

    private static boolean isAddedFlag(String v) {
        String t = v != null ? v.strip() : "";
        return ADDED_FLAG.equals(t) || "true".equalsIgnoreCase(t) || "追加".equals(t);
    }

    public static boolean isCombinationLockValue(String raw) {
        return isLockFlag(raw);
    }

    public static boolean isCombinationAddedValue(String raw) {
        return isAddedFlag(raw);
    }
    public static boolean[] visibilityMask(
            List<String> titles, int leadingCols, java.util.Set<String> focusNormalizedKeys) {
        int n = titles != null ? titles.size() : 0;
        boolean[] vis = new boolean[n];
        boolean focusing = focusNormalizedKeys != null && !focusNormalizedKeys.isEmpty();
        int lead = Math.max(0, leadingCols);
        for (int i = 0; i < n; i++) {
            if (i < lead) {
                vis[i] = true;
                continue;
            }
            String title = titles.get(i);
            String[] pm = splitEquipmentTitle(title);
            if (pm[0].isEmpty() && pm[1].isEmpty()) {
                vis[i] = false;
                continue;
            }
            if (!focusing) {
                vis[i] = true;
                continue;
            }
            String key = MasterTeamCombinationTableReader.normalizedComboKey(pm[0], pm[1]);
            vis[i] = !key.isEmpty() && focusNormalizedKeys.contains(key);
        }
        return vis;
    }

    /**
     * 「表示する設備を選ぶ」のチェック結果を工程+機械キーへ変換する。全設備を出しているときは空集合（絞り込みなし）。
     */
    public static Set<String> focusKeysFromVisibility(
            List<String> titles, int leadingCols, boolean[] visible) {
        LinkedHashSet<String> all = new LinkedHashSet<>();
        LinkedHashSet<String> selected = new LinkedHashSet<>();
        int n = titles != null ? titles.size() : 0;
        int lead = Math.max(0, leadingCols);
        for (int i = lead; i < n; i++) {
            String[] pm = splitEquipmentTitle(titles.get(i));
            if (pm[0].isEmpty() && pm[1].isEmpty()) {
                continue;
            }
            String key = MasterTeamCombinationTableReader.normalizedComboKey(pm[0], pm[1]);
            if (key.isEmpty()) {
                continue;
            }
            all.add(key);
            if (visible != null && i < visible.length && visible[i]) {
                selected.add(key);
            }
        }
        if (all.isEmpty() || selected.size() == all.size()) {
            return Set.of();
        }
        return Set.copyOf(selected);
    }

    /**
     * 組み合わせ表の本文行を表示するか。{@code focusKeys} が空ならすべて表示。空の追加入力行は残す。
     */
    public static boolean combinationDisplayRowVisible(
            List<String> header, List<String> dataRow, Set<String> focusKeys) {
        if (focusKeys == null || focusKeys.isEmpty()) {
            return true;
        }
        int procCol = headerIndex(header, "工程名");
        int machCol = headerIndex(header, "機械名");
        String proc =
                procCol >= 0 && dataRow != null && procCol < dataRow.size() && dataRow.get(procCol) != null
                        ? dataRow.get(procCol)
                        : "";
        String mach =
                machCol >= 0 && dataRow != null && machCol < dataRow.size() && dataRow.get(machCol) != null
                        ? dataRow.get(machCol)
                        : "";
        if (proc.isBlank() && mach.isBlank()) {
            return true;
        }
        String key = MasterTeamCombinationTableReader.normalizedComboKey(proc, mach);
        return !key.isEmpty() && focusKeys.contains(key);
    }

    /**
     * 組み合わせ表グリッドで隠す行。フィルタ行（{@code firstDataRow} 未満）は対象外。
     */
    public static BitSet combinationHiddenGridRows(
            List<List<String>> originalRows, int gridRowCount, int firstDataRow, Set<String> focusKeys) {
        BitSet hidden = new BitSet(Math.max(0, gridRowCount));
        if (focusKeys == null || focusKeys.isEmpty()) {
            return hidden;
        }
        List<String> header = headerRow(originalRows);
        List<List<String>> display = displayRows(SheetKind.COMBINATIONS, originalRows);
        int first = Math.max(0, firstDataRow);
        int n = Math.max(0, gridRowCount);
        for (int i = 0; i < display.size(); i++) {
            int gridRow = first + i;
            if (gridRow < 0 || gridRow >= n) {
                continue;
            }
            if (!combinationDisplayRowVisible(header, display.get(i), focusKeys)) {
                hidden.set(gridRow);
            }
        }
        return hidden;
    }

    public static boolean[] mandatoryLeadingMask(int colCount, int leadingCols) {
        int n = Math.max(0, colCount);
        boolean[] m = new boolean[n];
        int lead = Math.min(n, Math.max(0, leadingCols));
        for (int i = 0; i < lead; i++) {
            m[i] = true;
        }
        return m;
    }

    public static String dialogColumnLabel(String title) {
        String t = title != null ? title.strip() : "";
        if (t.isEmpty()) {
            return "（空列）";
        }
        int nl = t.indexOf('\n');
        if (nl < 0) {
            return t;
        }
        String a = t.substring(0, nl).strip();
        String b = t.substring(nl + 1).strip();
        if (a.isEmpty()) {
            return b;
        }
        if (b.isEmpty()) {
            return a;
        }
        return a + " / " + b;
    }

    public static List<String> dialogColumnLabels(List<String> titles) {
        if (titles == null || titles.isEmpty()) {
            return List.of();
        }
        List<String> out = new ArrayList<>(titles.size());
        for (String t : titles) {
            out.add(dialogColumnLabel(t));
        }
        return List.copyOf(out);
    }

    public static List<List<String>> restoreTitleRows(
            SheetKind kind, List<String> titles, List<List<String>> displayRows) {
        List<List<String>> body = displayRows != null ? displayRows : List.of();
        if (kind == SheetKind.COMBINATIONS) {
            if (!body.isEmpty() && isColumnTitleSourceRow(kind, 0, body)) {
                return copyRows(body);
            }
            if (titles == null || titles.isEmpty()) {
                return copyRows(body);
            }
            List<List<String>> out = new ArrayList<>();
            out.add(paddedRow(titles, titles.size()));
            out.addAll(mutableCopy(body));
            return freeze(out);
        }
        if (!body.isEmpty() && "工程名".equals(cell(body, 0, 0))) {
            return copyRows(body);
        }
        int cols = Math.max(widthOf(body), titles != null ? titles.size() : 1);
        int firstEq = kind == SheetKind.NEED || kind == SheetKind.SPEED ? 3 : 1;
        List<String> proc = new ArrayList<>(cols);
        List<String> mach = new ArrayList<>(cols);
        proc.add("工程名");
        mach.add("機械名");
        for (int c = 1; c < cols; c++) {
            if (c < firstEq) {
                proc.add("");
                mach.add("");
                continue;
            }
            String title = titles != null && c < titles.size() ? titles.get(c) : "";
            String[] pm = splitEquipmentTitle(title);
            proc.add(pm[0]);
            mach.add(pm[1]);
        }
        List<List<String>> out = new ArrayList<>();
        out.add(proc);
        out.add(mach);
        out.addAll(mutableCopy(body));
        return freeze(out);
    }

    public static List<Double> preferredColumnWidths(
            List<List<String>> dataRows, int colCount, List<String> titles) {
        List<Double> fromData = preferredColumnWidths(dataRows, colCount);
        if (titles == null || titles.isEmpty()) {
            return fromData;
        }
        List<Double> out = new ArrayList<>(fromData);
        for (int c = 0; c < out.size() && c < titles.size(); c++) {
            double tw = measureTextWidth(titles.get(c));
            if (tw > 28.0) {
                out.set(c, Math.min(COL_WIDTH_MAX, Math.max(out.get(c), Math.max(COL_WIDTH_MIN, tw))));
            }
        }
        return List.copyOf(out);
    }

    static String equipmentColumnTitle(String process, String machine) {
        String p = process != null ? process.strip() : "";
        String m = machine != null ? machine.strip() : "";
        if (!p.isEmpty() && !m.isEmpty()) {
            return p + "\n" + m;
        }
        if (!p.isEmpty()) {
            return p;
        }
        return m;
    }

    static String[] splitEquipmentTitle(String title) {
        String t = title != null ? title.strip() : "";
        if (t.isEmpty()) {
            return new String[] {"", ""};
        }
        int nl = t.indexOf('\n');
        if (nl < 0) {
            return new String[] {t, ""};
        }
        return new String[] {t.substring(0, nl).strip(), t.substring(nl + 1).strip()};
    }

    private static List<String> paddedRow(List<String> src, int cols) {
        List<String> row = new ArrayList<>(cols);
        for (int c = 0; c < cols; c++) {
            String v = src != null && c < src.size() && src.get(c) != null ? src.get(c) : "";
            row.add(v);
        }
        return row;
    }

    public static String comboRowStyle(String process, String machine) {
        return combinationRowStyle(process, machine, "", false, false);
    }

    public static String comboRowStyle(String process, String machine, String comboCell) {
        return combinationRowStyle(process, machine, comboCell, false, false);
    }

    public static String combinationRowStyle(
            String process, String machine, String comboCell, boolean added, boolean locked) {
        if (added) {
            return "-fx-background-color: "
                    + ADDED_ROW_BG
                    + "; -fx-control-inner-background: "
                    + ADDED_ROW_BG
                    + "; -fx-text-fill: #111111;";
        }
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

    public static boolean isColumnTitleSourceRow(
            SheetKind kind, int dataRow, List<List<String>> rows) {
        if (kind == SheetKind.COMBINATIONS) {
            String a = cell(rows, dataRow, 0);
            return "組み合わせ行ID".equals(a) || "組合せ行ID".equals(a) || "インデックス".equals(a);
        }
        String a = cell(rows, dataRow, 0);
        if (isProcessOrMachineHeaderLabel(a)) {
            return true;
        }
        return kind == SheetKind.NEED && isNeedColumnCaptionRow(rows, dataRow);
    }

    public static TitleRowKind titleRowKind(SheetKind kind, int dataRow, List<List<String>> rows) {
        if (!isColumnTitleSourceRow(kind, dataRow, rows)) {
            return TitleRowKind.NONE;
        }
        String a = cell(rows, dataRow, 0);
        if (a.startsWith("機械名")) {
            return TitleRowKind.MACHINE;
        }
        if (a.startsWith("工程名")) {
            return TitleRowKind.PROCESS;
        }
        return TitleRowKind.OTHER;
    }

    static boolean isProcessOrMachineHeaderLabel(String a) {
        return a.startsWith("工程名") || a.startsWith("機械名");
    }

    static boolean isNeedColumnCaptionRow(List<List<String>> rows, int dataRow) {
        String a = cell(rows, dataRow, 0);
        String b = cell(rows, dataRow, 1);
        String c = cell(rows, dataRow, 2);
        return a.isEmpty() && "依頼NO条件".equals(b) && "備考".equals(c);
    }

    public static boolean isEditable(SheetKind kind, int dataRow, int col, List<List<String>> rows) {
        return isEditable(kind, dataRow, col, rows, rows);
    }

    public static boolean isEditable(
            SheetKind kind,
            int dataRow,
            int col,
            List<List<String>> displayRows,
            List<List<String>> originalRows) {
        if (kind == null || dataRow < 0 || col < 0) {
            return false;
        }
        List<List<String>> display = displayRows != null ? displayRows : List.of();
        List<List<String>> original = originalRows != null ? originalRows : display;
        return switch (kind) {
            case SKILLS -> isSkillsEditable(dataRow, col, display);
            case NEED -> isNeedEditable(dataRow, col, display);
            case SPEED -> isSpeedEditable(dataRow, col, display);
            case COMBINATIONS -> isCombinationsEditable(dataRow, col, display, original);
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
            case SKILLS -> isSkillsSkillValueCell(dataRow, col, rows) && parseOpAs(v) == null;
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

    private static boolean isSkillsEditable(int dataRow, int col, List<List<String>> rows) {
        return col != 0 || !isColumnTitleSourceRow(SheetKind.SKILLS, dataRow, rows);
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

    private static boolean isCombinationsEditable(
            int dataRow, int col, List<List<String>> display, List<List<String>> original) {
        if (isColumnTitleSourceRow(SheetKind.COMBINATIONS, dataRow, display)) {
            return false;
        }
        List<String> header = headerRow(original);
        int comboCol = headerIndex(header, "工程+機械", "工程＋機械");
        if (comboCol >= 0 && col == comboCol) {
            return false;
        }
        if (headerIndex(header, COL_ADDED_ROW) == col) {
            return false;
        }
        if (isCombinationLockColumn(header, col)) {
            return true;
        }
        int lockCol = headerIndex(header, COL_EDIT_LOCK);
        return lockCol < 0 || !isLockFlag(cell(display, dataRow, lockCol));
    }

    private static boolean isNeedStructureLabel(String a) {
        if (a.isEmpty()) {
            return false;
        }
        return isProcessOrMachineHeaderLabel(a)
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
        return isProcessOrMachineHeaderLabel(a)
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
        return findFirstRowExact(rows, "工程名") == 0 && (dataRow == 3 || dataRow == 4);
    }

    private static boolean isCombinationsNumericInvalid(
            int dataRow, int col, List<List<String>> rows, String v) {
        if (isColumnTitleSourceRow(SheetKind.COMBINATIONS, dataRow, rows)) {
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

    private static int findFirstRowExact(List<List<String>> rows, String label) {
        if (rows == null || label == null) {
            return -1;
        }
        for (int r = 0; r < rows.size(); r++) {
            if (label.equals(cell(rows, r, 0))) {
                return r;
            }
        }
        return -1;
    }

    private static int findProcessHeaderRow(List<List<String>> rows) {
        return findHeaderRowPreferExact(rows, "工程名");
    }

    private static int findMachineHeaderRow(List<List<String>> rows) {
        return findHeaderRowPreferExact(rows, "機械名");
    }

    private static void ensureProcessMachineHeaderRows(SheetKind kind, List<List<String>> rows) {
        int firstEq = kind == SheetKind.NEED || kind == SheetKind.SPEED ? 3 : 1;
        int minWidth = Math.max(widthOf(rows), firstEq);
        if (findProcessHeaderRow(rows) < 0) {
            List<String> proc = new ArrayList<>();
            proc.add("工程名");
            while (proc.size() < minWidth) {
                proc.add("");
            }
            rows.add(0, proc);
        }
        if (findMachineHeaderRow(rows) < 0) {
            int procRow = findProcessHeaderRow(rows);
            int insertAt = procRow >= 0 ? procRow + 1 : 0;
            List<String> mach = new ArrayList<>();
            mach.add("機械名");
            while (mach.size() < minWidth) {
                mach.add("");
            }
            rows.add(insertAt, mach);
        }
        padWidth(rows, minWidth);
    }

    private static void padWidth(List<List<String>> rows, int cols) {
        for (List<String> row : rows) {
            while (row.size() < cols) {
                row.add("");
            }
        }
    }

    private static void setCell(List<List<String>> rows, int r, int c, String value) {
        while (rows.size() <= r) {
            rows.add(new ArrayList<>());
        }
        List<String> row = rows.get(r);
        while (row.size() <= c) {
            row.add("");
        }
        row.set(c, value != null ? value : "");
    }

    private static int findHeaderRowPreferExact(List<List<String>> rows, String exact) {
        int alias = -1;
        if (rows == null || exact == null) {
            return -1;
        }
        for (int r = 0; r < rows.size(); r++) {
            String a = cell(rows, r, 0);
            if (exact.equals(a)) {
                return r;
            }
            if (alias < 0 && a.startsWith(exact)) {
                alias = r;
            }
        }
        return alias;
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
        double best = 0;
        for (String line : raw.split("\n", -1)) {
            double w = 28.0;
            for (int i = 0; i < line.length(); i++) {
                w += line.charAt(i) > 127 ? 14.0 : 8.0;
            }
            best = Math.max(best, w);
        }
        return best;
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
