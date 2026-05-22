package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashSet;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup;

/**
 * {@code master.xls(m)} の「組み合わせ表」シートから工程+機械キー（{@code 工程名+機械名} 形式・正規化済み）を読む。
 * Python {@code load_team_combination_presets_from_master} と同趣旨。
 */
public final class MasterTeamCombinationTableReader {

    public static final String SHEET_NAME = "組み合わせ表";

    private MasterTeamCombinationTableReader() {}

    /** 正規化済み {@code proc+mach} キー集合。シートが無い／空のときは空集合。 */
    public static Set<String> readNormalizedComboKeys(Path workbookPath) throws IOException {
        Objects.requireNonNull(workbookPath, "workbookPath");
        if (!Files.isRegularFile(workbookPath)) {
            throw new IOException("not a file: " + workbookPath);
        }
        DataFormatter fmt = new DataFormatter(Locale.ROOT);
        LinkedHashSet<String> keys = new LinkedHashSet<>();
        try (InputStream in = Files.newInputStream(workbookPath);
                Workbook wb = WorkbookFactory.create(in)) {
            Sheet sh = wb.getSheet(SHEET_NAME);
            if (sh == null) {
                return Set.of();
            }
            Row header = sh.getRow(0);
            if (header == null) {
                return Set.of();
            }
            int lastCol = header.getLastCellNum();
            int procCol = -1;
            int machCol = -1;
            int comboCol = -1;
            for (int c = 0; c < lastCol; c++) {
                String h = normHeader(fmt, header.getCell(c));
                if ("工程名".equals(h)) {
                    procCol = c;
                } else if ("機械名".equals(h)) {
                    machCol = c;
                } else if ("工程+機械".equals(h) || "工程＋機械".equals(h)) {
                    comboCol = c;
                }
            }
            int lastRow = sh.getLastRowNum();
            for (int r = 1; r <= lastRow; r++) {
                Row row = sh.getRow(r);
                if (row == null) {
                    continue;
                }
                String proc = procCol >= 0 ? cellStr(fmt, row.getCell(procCol)) : "";
                String mach = machCol >= 0 ? cellStr(fmt, row.getCell(machCol)) : "";
                String comboCell = comboCol >= 0 ? cellStr(fmt, row.getCell(comboCol)) : "";
                String key = comboKeyFromCells(proc, mach, comboCell);
                if (!key.isEmpty()) {
                    keys.add(key);
                }
            }
        }
        return Set.copyOf(keys);
    }

    public static String comboKeyFromCells(String process, String machine, String comboCell) {
        String proc = process != null ? process.strip() : "";
        String mach = machine != null ? machine.strip() : "";
        if (!proc.isEmpty() && !mach.isEmpty()) {
            return normalizedComboKey(proc, mach);
        }
        String combo = comboCell != null ? comboCell.strip() : "";
        if (combo.isEmpty()) {
            return "";
        }
        int plus = combo.indexOf('+');
        if (plus > 0 && plus < combo.length() - 1) {
            return normalizedComboKey(combo.substring(0, plus), combo.substring(plus + 1));
        }
        return normalizeComboKeyLiteral(combo);
    }

    public static String normalizedComboKey(String process, String machine) {
        String p = AladdinShapedPlanQtyLookup.normalizeProcessNameForRuleMatch(process);
        String m = normalizeEquipmentMatchKey(machine);
        if (p.isEmpty() || m.isEmpty()) {
            return "";
        }
        return p + "+" + m;
    }

    private static String normalizeEquipmentMatchKey(String val) {
        if (val == null || val.isBlank()) {
            return "";
        }
        String t = java.text.Normalizer.normalize(val, java.text.Normalizer.Form.NFKC);
        t = t.replace('\u00a0', ' ').replace('\u3000', ' ');
        t = t.replaceAll("[\u200b\u200c\u200d\ufeff]", "");
        return t.replaceAll("\\s+", " ").strip();
    }

    private static String normalizeComboKeyLiteral(String combo) {
        String c = combo.strip();
        if (c.isEmpty()) {
            return "";
        }
        int plus = c.indexOf('+');
        if (plus > 0 && plus < c.length() - 1) {
            return normalizedComboKey(c.substring(0, plus), c.substring(plus + 1));
        }
        return AladdinShapedPlanQtyLookup.normalizeProcessNameForRuleMatch(c);
    }

    private static String normHeader(DataFormatter fmt, Cell cell) {
        String s = cellStr(fmt, cell);
        if (s.isEmpty()) {
            return "";
        }
        return s.replace("組合せ", "組み合わせ");
    }

    private static String cellStr(DataFormatter fmt, Cell cell) {
        if (cell == null) {
            return "";
        }
        return ExcelCellReadSupport.normalizeCommaDigitArtifacts(fmt.formatCellValue(cell).trim());
    }
}
