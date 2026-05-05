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

    private static boolean detectTwoHeaderRows(Sheet sh, DataFormatter fmt, int lastRow) {
        if (lastRow < 2) {
            return false;
        }
        Row r0 = sh.getRow(0);
        Row r1 = sh.getRow(1);
        if (r0 == null || r1 == null) {
            return false;
        }
        int nonEmpty = 0;
        int maxC = Math.max(r0.getLastCellNum(), r1.getLastCellNum());
        for (int c = 1; c < maxC; c++) {
            String p = cellStr(fmt, r0.getCell(c));
            String m = cellStr(fmt, r1.getCell(c));
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
        return fmt.formatCellValue(cell).trim();
    }
}
