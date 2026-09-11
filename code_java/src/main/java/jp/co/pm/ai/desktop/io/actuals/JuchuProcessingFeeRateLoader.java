package jp.co.pm.ai.desktop.io.actuals;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import jp.co.pm.ai.desktop.reconciliation.JuchuSheetColumnLayout;

/**
 * 受注ﾌｧｲﾙシートから依頼No → 加工賃情報（AH・AO・加工内容）を読む。
 *
 * <p>列位置は index 固定（AH={@link JuchuSheetColumnLayout.Col#KAKOCHIN}、AO、Z=加工内容）。
 * 同一依頼No は後勝ち。AH の改行複数値は <b>末尾</b>の数値（最終工程単価）を採用する。
 * AO は依頼NOごとの加工賃合計（円）。
 */
public final class JuchuProcessingFeeRateLoader {

    public static final String SHEET_NAME = "受注ﾌｧｲﾙ";
    /** 見出し行（1-based）。工場既定に合わせる。 */
    public static final int DEFAULT_HEADER_ROW_ONE_BASED = 3;

    /** 受注ﾌｧｲﾙ AO 列（0-based）。依頼NOごとの加工賃合計。 */
    public static final int AO_COLUMN_INDEX = JuchuSheetColumnLayout.columnLetterToIndex("AO");

    private static final DataFormatter FORMATTER = new DataFormatter(Locale.JAPAN);

    private JuchuProcessingFeeRateLoader() {}

    /**
     * @param rateAhYenPerM AH 末尾行の単価（円/m）。欠落時は {@code null}
     * @param totalAoYen AO の加工賃合計（円）。欠落・非数値は {@code null}
     * @param processContent Z 列の加工内容（カンマ区切り工程列）
     */
    public record FeeInfo(Double rateAhYenPerM, Double totalAoYen, String processContent) {
        public FeeInfo {
            processContent = processContent == null ? "" : processContent.strip();
        }

        public boolean hasAo() {
            return totalAoYen != null && totalAoYen > 0 && !Double.isNaN(totalAoYen);
        }

        public boolean hasAh() {
            return rateAhYenPerM != null && !Double.isNaN(rateAhYenPerM);
        }
    }

    /**
     * @param workbookPath 受注 Excel（.xlsx / .xlsm）
     * @return 依頼No（trim）→ 加工賃単価（AH）。空マップもあり得る
     */
    public static Map<String, Double> loadRates(Path workbookPath) throws IOException {
        Map<String, Double> out = new LinkedHashMap<>();
        for (Map.Entry<String, FeeInfo> e : loadFeeInfo(workbookPath).entrySet()) {
            if (e.getValue().hasAh()) {
                out.put(e.getKey(), e.getValue().rateAhYenPerM());
            }
        }
        return Collections.unmodifiableMap(out);
    }

    public static Map<String, Double> loadRates(Path workbookPath, int headerRowOneBased)
            throws IOException {
        Map<String, Double> out = new LinkedHashMap<>();
        for (Map.Entry<String, FeeInfo> e : loadFeeInfo(workbookPath, headerRowOneBased).entrySet()) {
            if (e.getValue().hasAh()) {
                out.put(e.getKey(), e.getValue().rateAhYenPerM());
            }
        }
        return Collections.unmodifiableMap(out);
    }

    public static Map<String, Double> loadRates(Workbook wb, int headerRowOneBased) {
        Map<String, Double> out = new LinkedHashMap<>();
        for (Map.Entry<String, FeeInfo> e : loadFeeInfo(wb, headerRowOneBased).entrySet()) {
            if (e.getValue().hasAh()) {
                out.put(e.getKey(), e.getValue().rateAhYenPerM());
            }
        }
        return Collections.unmodifiableMap(out);
    }

    public static Map<String, FeeInfo> loadFeeInfo(Path workbookPath) throws IOException {
        return loadFeeInfo(workbookPath, DEFAULT_HEADER_ROW_ONE_BASED);
    }

    public static Map<String, FeeInfo> loadFeeInfo(Path workbookPath, int headerRowOneBased)
            throws IOException {
        Objects.requireNonNull(workbookPath, "workbookPath");
        if (!Files.isRegularFile(workbookPath)) {
            throw new IOException("受注ファイルがありません: " + workbookPath);
        }
        try (InputStream in = Files.newInputStream(workbookPath);
                Workbook wb = WorkbookFactory.create(in)) {
            return loadFeeInfo(wb, headerRowOneBased);
        }
    }

    public static Map<String, FeeInfo> loadFeeInfo(Workbook wb, int headerRowOneBased) {
        Objects.requireNonNull(wb, "wb");
        if (headerRowOneBased < 1) {
            throw new IllegalArgumentException("headerRowOneBased must be >= 1");
        }
        Sheet sheet = wb.getSheet(SHEET_NAME);
        if (sheet == null) {
            return Map.of();
        }
        FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();
        int iraiCol = JuchuSheetColumnLayout.Col.IRAI_NO.columnIndex();
        int feeCol = JuchuSheetColumnLayout.Col.KAKOCHIN.columnIndex();
        int kakoCol = JuchuSheetColumnLayout.Col.KAKO_NAIYO.columnIndex();
        int aoCol = AO_COLUMN_INDEX;
        int firstDataRow = headerRowOneBased; // 0-based: header is headerRowOneBased-1
        Map<String, FeeInfo> out = new LinkedHashMap<>();
        int last = Math.min(sheet.getLastRowNum(), firstDataRow + 50_000);
        for (int r = firstDataRow; r <= last; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            String irai = cellText(row.getCell(iraiCol), eval).strip();
            if (irai.isEmpty() || "0".equals(irai)) {
                continue;
            }
            Double rate = parseFeeRateLastLine(cellText(row.getCell(feeCol), eval));
            Double ao = parseAoYen(row.getCell(aoCol), eval);
            String kako = cellText(row.getCell(kakoCol), eval);
            if (rate == null && ao == null && kako.isBlank()) {
                continue;
            }
            out.put(irai, new FeeInfo(rate, ao, kako));
        }
        return Collections.unmodifiableMap(out);
    }

    /** 互換: 先頭行の数値。空・非数値は null。 */
    static Double parseFeeRate(String raw) {
        return parseFeeRateLine(raw, false);
    }

    /** AH 複数行は末尾の数値（最終工程単価）。 */
    static Double parseFeeRateLastLine(String raw) {
        return parseFeeRateLine(raw, true);
    }

    private static Double parseFeeRateLine(String raw, boolean lastLine) {
        if (raw == null) {
            return null;
        }
        String s = raw.strip();
        if (s.isEmpty()) {
            return null;
        }
        String[] lines = s.split("\\R");
        String pick = null;
        if (lastLine) {
            for (int i = lines.length - 1; i >= 0; i--) {
                String t = lines[i].strip();
                if (!t.isEmpty()) {
                    pick = t;
                    break;
                }
            }
        } else {
            pick = lines[0].strip();
        }
        if (pick == null || pick.isEmpty()) {
            return null;
        }
        return parsePlainNumber(pick);
    }

    static Double parseAoYen(Cell cell, FormulaEvaluator eval) {
        if (cell == null) {
            return null;
        }
        CellType type = cell.getCellType();
        if (type == CellType.FORMULA && eval != null) {
            try {
                type = eval.evaluateFormulaCell(cell);
            } catch (RuntimeException ignored) {
                return parsePlainNumber(FORMATTER.formatCellValue(cell));
            }
        }
        if (type == CellType.NUMERIC) {
            double n = cell.getNumericCellValue();
            if (Double.isNaN(n) || Double.isInfinite(n)) {
                return null;
            }
            return n;
        }
        return parsePlainNumber(FORMATTER.formatCellValue(cell));
    }

    private static Double parsePlainNumber(String raw) {
        if (raw == null) {
            return null;
        }
        String s = raw.strip();
        if (s.isEmpty()) {
            return null;
        }
        s = s.replace(",", "").replace("，", "").replace("¥", "").replace("円", "").strip();
        if (s.isEmpty()) {
            return null;
        }
        try {
            double v = Double.parseDouble(s);
            if (Double.isNaN(v) || Double.isInfinite(v)) {
                return null;
            }
            return v;
        } catch (NumberFormatException ex) {
            return null;
        }
    }

    private static String cellText(Cell cell, FormulaEvaluator eval) {
        if (cell == null) {
            return "";
        }
        CellType type = cell.getCellType();
        if (type == CellType.FORMULA && eval != null) {
            try {
                type = eval.evaluateFormulaCell(cell);
            } catch (RuntimeException ignored) {
                return FORMATTER.formatCellValue(cell);
            }
        }
        if (type == CellType.NUMERIC) {
            double n = cell.getNumericCellValue();
            if (n == Math.rint(n) && Math.abs(n) < 1e15) {
                return Long.toString((long) Math.rint(n));
            }
            return Double.toString(n);
        }
        return FORMATTER.formatCellValue(cell).strip();
    }
}
