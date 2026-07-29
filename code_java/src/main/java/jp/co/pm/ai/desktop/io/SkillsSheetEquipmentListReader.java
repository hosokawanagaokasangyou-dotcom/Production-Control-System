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

import jp.co.pm.ai.desktop.io.MasterTeamCombinationTableReader;

/**
 * {@code master.xls(x/m)} の {@code skills} シートから、Python {@code load_skills_and_needs} の 2 段ヘッダ分岐と同趣旨の
 * 「工程名+機械名」列キー一覧を出現順で返す（重複列は先勝ち）。
 */
public final class SkillsSheetEquipmentListReader {

    private SkillsSheetEquipmentListReader() {}

    public static List<String> readEquipmentProcPlusMachineCombos(Path workbookPath) throws IOException {
        Objects.requireNonNull(workbookPath, "workbookPath");
        if (!Files.isRegularFile(workbookPath)) {
            throw new IOException("not a file: " + workbookPath);
        }
        DataFormatter fmt = new DataFormatter(Locale.ROOT);
        try (InputStream in = Files.newInputStream(workbookPath);
                Workbook wb = WorkbookFactory.create(in)) {
            Sheet sh = wb.getSheet("skills");
            if (sh == null) {
                return List.of();
            }
            int lastRow = sh.getLastRowNum();
            if (lastRow < 2) {
                return List.of();
            }
            Row r0 = sh.getRow(0);
            Row r1 = sh.getRow(1);
            if (r0 == null || r1 == null) {
                return List.of();
            }
            int maxC = Math.max(r0.getLastCellNum(), r1.getLastCellNum());
            int nonEmptyPm = 0;
            for (int c = 1; c < maxC; c++) {
                String p = cellStr(fmt, r0.getCell(c));
                String m = cellStr(fmt, r1.getCell(c));
                if (!p.isEmpty()
                        && !m.isEmpty()
                        && !"nan".equalsIgnoreCase(p)
                        && !"nan".equalsIgnoreCase(m)) {
                    nonEmptyPm++;
                }
            }
            if (nonEmptyPm <= 0) {
                return List.of();
            }
            LinkedHashSet<String> seen = new LinkedHashSet<>();
            List<String> out = new ArrayList<>();
            for (int c = 1; c < maxC; c++) {
                String p = cellStr(fmt, r0.getCell(c));
                String m = cellStr(fmt, r1.getCell(c));
                if (p.isEmpty()
                        || m.isEmpty()
                        || "nan".equalsIgnoreCase(p)
                        || "nan".equalsIgnoreCase(m)) {
                    continue;
                }
                String combo = p + "+" + m;
                if (seen.add(combo)) {
                    out.add(combo);
                }
            }
            return List.copyOf(out);
        }
    }

    /**
     * skills シート 2 段ヘッダの「工程名+機械名」列を正規化キー集合で返す。
     * 1 行ヘッダ形式・シート無しのときは空集合。
     */
    public static Set<String> readNormalizedComboKeys(Path workbookPath) throws IOException {
        List<String> combos = readEquipmentProcPlusMachineCombos(workbookPath);
        LinkedHashSet<String> keys = new LinkedHashSet<>();
        for (String combo : combos) {
            int plus = combo.indexOf('+');
            if (plus <= 0 || plus >= combo.length() - 1) {
                continue;
            }
            String key =
                    MasterTeamCombinationTableReader.normalizedComboKey(
                            combo.substring(0, plus), combo.substring(plus + 1));
            if (!key.isEmpty()) {
                keys.add(key);
            }
        }
        return Set.copyOf(keys);
    }

    private static String cellStr(DataFormatter fmt, Cell cell) {
        if (cell == null) {
            return "";
        }
        return ExcelCellReadSupport.normalizeCommaDigitArtifacts(fmt.formatCellValue(cell).trim());
    }
}
