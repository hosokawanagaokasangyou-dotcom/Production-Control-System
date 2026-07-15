package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * {@code master.xls(x/m)} の {@code skills} シートからメンバー名一覧を読む。
 *
 * <p>Python {@code planning_core._core} の skills 読込（2段ヘッダ／1行ヘッダ旧形式）に概ね整合する。
 */
public final class SkillsSheetMemberReader {

    private static final Pattern OP_AS_SKILL = Pattern.compile("^(OP|AS)\\d*$", Pattern.CASE_INSENSITIVE);

    private static final Set<String> MEMBER_HEADER_NAMES =
            Set.of(
                    "\u30e1\u30f3\u30d0\u30fc",
                    "\u62c5\u5f53\u8005",
                    "\u4e26\u3073",
                    "\u4f5c\u696d\u8005");

    private SkillsSheetMemberReader() {}

    /**
     * skills シートのメンバー表示名を出現順（重複は後続を捨てる）で返す。
     *
     * @throws IOException ファイルやシートが無い場合
     */
    public static List<String> readMemberDisplayNames(Path workbookPath) throws IOException {
        Objects.requireNonNull(workbookPath, "workbookPath");
        if (!Files.isRegularFile(workbookPath)) {
            throw new IOException("not a file: " + workbookPath);
        }
        DataFormatter fmt = new DataFormatter(Locale.ROOT);
        try (InputStream in = Files.newInputStream(workbookPath);
                Workbook wb = WorkbookFactory.create(in)) {
            Sheet sh = wb.getSheet("skills");
            if (sh == null) {
                throw new IOException("sheet 'skills' not found in " + workbookPath);
            }
            int lastRow = sh.getLastRowNum();
            if (lastRow < 0) {
                return List.of();
            }
            boolean twoHeader = detectTwoHeaderRows(sh, fmt, lastRow);
            List<String> raw =
                    twoHeader
                            ? readMembersTwoHeader(sh, fmt, lastRow)
                            : readMembersSingleHeader(sh, fmt, lastRow);
            LinkedHashSet<String> seen = new LinkedHashSet<>();
            List<String> out = new ArrayList<>();
            for (String s : raw) {
                if (s != null && !s.isBlank()) {
                    String t = s.strip();
                    if (seen.add(t)) {
                        out.add(t);
                    }
                }
            }
            return List.copyOf(out);
        }
    }

    /**
     * 対象の工程名＋機械名に {@code OP} または {@code AS} 資格を持つメンバーを skills 行順で返す。
     *
     * <p>資格セルは Python {@code parse_op_as_skill_cell} と同様、空白を除いた
     * {@code OP}/{@code AS} と任意の優先度整数だけを資格ありと解釈する。
     */
    public static List<String> readQualifiedMemberDisplayNames(
            Path workbookPath, String processName, String machineName) throws IOException {
        Objects.requireNonNull(workbookPath, "workbookPath");
        if (!Files.isRegularFile(workbookPath)) {
            throw new IOException("not a file: " + workbookPath);
        }
        String process = processName != null ? processName.strip() : "";
        String machine = machineName != null ? machineName.strip() : "";
        if (process.isEmpty() || machine.isEmpty()) {
            return List.of();
        }

        DataFormatter fmt = new DataFormatter(Locale.ROOT);
        try (InputStream in = Files.newInputStream(workbookPath);
                Workbook wb = WorkbookFactory.create(in)) {
            Sheet sh = wb.getSheet("skills");
            if (sh == null) {
                throw new IOException("sheet 'skills' not found in " + workbookPath);
            }
            int lastRow = sh.getLastRowNum();
            if (lastRow < 0) {
                return List.of();
            }
            return detectTwoHeaderRows(sh, fmt, lastRow)
                    ? readQualifiedTwoHeader(sh, fmt, lastRow, process, machine)
                    : readQualifiedSingleHeader(sh, fmt, lastRow, process, machine);
        }
    }

    private static List<String> readQualifiedTwoHeader(
            Sheet sh, DataFormatter fmt, int lastRow, String process, String machine) {
        Row processRow = sh.getRow(0);
        Row machineRow = sh.getRow(1);
        int maxC = Math.max(processRow.getLastCellNum(), machineRow.getLastCellNum());
        List<Integer> skillColumns = new ArrayList<>();
        for (int c = 1; c < maxC; c++) {
            if (process.equals(cellStr(fmt, processRow.getCell(c)))
                    && machine.equals(cellStr(fmt, machineRow.getCell(c)))) {
                skillColumns.add(c);
            }
        }
        return qualifiedNamesFromRows(sh, fmt, 2, lastRow, 0, skillColumns);
    }

    private static List<String> readQualifiedSingleHeader(
            Sheet sh, DataFormatter fmt, int lastRow, String process, String machine) {
        Row header = sh.getRow(0);
        if (header == null) {
            return List.of();
        }
        int memberCol = -1;
        int comboCol = -1;
        int machineCol = -1;
        int processCol = -1;
        String combo = process + "+" + machine;
        for (int c = 0; c < header.getLastCellNum(); c++) {
            String value = cellStr(fmt, header.getCell(c));
            if (memberCol < 0 && MEMBER_HEADER_NAMES.contains(value)) {
                memberCol = c;
            }
            if (comboCol < 0 && combo.equals(value)) {
                comboCol = c;
            }
            if (machineCol < 0 && machine.equals(value)) {
                machineCol = c;
            }
            if (processCol < 0 && process.equals(value)) {
                processCol = c;
            }
        }
        if (memberCol < 0) {
            memberCol = 0;
        }
        int skillCol = comboCol >= 0 ? comboCol : machineCol >= 0 ? machineCol : processCol;
        return skillCol >= 0
                ? qualifiedNamesFromRows(sh, fmt, 1, lastRow, memberCol, List.of(skillCol))
                : List.of();
    }

    private static List<String> qualifiedNamesFromRows(
            Sheet sh,
            DataFormatter fmt,
            int firstRow,
            int lastRow,
            int memberCol,
            List<Integer> skillColumns) {
        LinkedHashSet<String> names = new LinkedHashSet<>();
        for (int r = firstRow; r <= lastRow; r++) {
            Row row = sh.getRow(r);
            if (row == null) {
                continue;
            }
            String name = cellStr(fmt, row.getCell(memberCol));
            if (name.isEmpty() || "nan".equalsIgnoreCase(name)) {
                continue;
            }
            for (int skillCol : skillColumns) {
                String skill = cellStr(fmt, row.getCell(skillCol)).replaceAll("\\s+", "");
                if (OP_AS_SKILL.matcher(skill).matches()) {
                    names.add(name.strip());
                    break;
                }
            }
        }
        return List.copyOf(names);
    }

    private static boolean detectTwoHeaderRows(Sheet sh, DataFormatter fmt, int lastRow) {
        if (lastRow < 2) {
            return false;
        }
        Row r0 = sh.getRow(0);
        Row r1 = sh.getRow(1);
        if (r0 == null || r1 == null) {
            return false;
        }
        String topLeft = cellStr(fmt, r0.getCell(0)).strip();
        String secondLeft = cellStr(fmt, r1.getCell(0)).strip();
        if (MEMBER_HEADER_NAMES.contains(topLeft)) {
            return false;
        }
        boolean structuralHeaderLabels =
                (topLeft.isEmpty() || "工程名".equals(topLeft))
                        && (secondLeft.isEmpty() || "機械名".equals(secondLeft));
        if (!structuralHeaderLabels) {
            return false;
        }
        int nonEmpty = 0;
        int maxC = Math.max(r0.getLastCellNum(), r1.getLastCellNum());
        for (int c = 1; c < maxC; c++) {
            String p = cellStr(fmt, r0.getCell(c));
            String m = cellStr(fmt, r1.getCell(c));
            if (OP_AS_SKILL.matcher(m.replaceAll("\\s+", "")).matches()) {
                return false;
            }
            if (!p.isEmpty()
                    && !m.isEmpty()
                    && !"nan".equalsIgnoreCase(p)
                    && !"nan".equalsIgnoreCase(m)) {
                nonEmpty++;
            }
        }
        return nonEmpty > 0;
    }

    private static List<String> readMembersTwoHeader(Sheet sh, DataFormatter fmt, int lastRow) {
        List<String> names = new ArrayList<>();
        for (int r = 2; r <= lastRow; r++) {
            Row row = sh.getRow(r);
            if (row == null) {
                continue;
            }
            String name = cellStr(fmt, row.getCell(0));
            if (name.isEmpty() || "nan".equalsIgnoreCase(name)) {
                continue;
            }
            names.add(name);
        }
        return names;
    }

    private static List<String> readMembersSingleHeader(Sheet sh, DataFormatter fmt, int lastRow) {
        Row head = sh.getRow(0);
        if (head == null) {
            return List.of();
        }
        int memberCol = -1;
        int maxC = head.getLastCellNum();
        for (int c = 0; c < maxC; c++) {
            String h = cellStr(fmt, head.getCell(c));
            if (MEMBER_HEADER_NAMES.contains(h.strip())) {
                memberCol = c;
                break;
            }
        }
        if (memberCol < 0) {
            memberCol = 0;
        }
        List<String> names = new ArrayList<>();
        for (int r = 1; r <= lastRow; r++) {
            Row row = sh.getRow(r);
            if (row == null) {
                continue;
            }
            String name = cellStr(fmt, row.getCell(memberCol));
            if (name.isEmpty() || "nan".equalsIgnoreCase(name)) {
                continue;
            }
            names.add(name);
        }
        return names;
    }

    private static String cellStr(DataFormatter fmt, Cell cell) {
        if (cell == null) {
            return "";
        }
        return ExcelCellReadSupport.normalizeCommaDigitArtifacts(fmt.formatCellValue(cell).trim());
    }
}
