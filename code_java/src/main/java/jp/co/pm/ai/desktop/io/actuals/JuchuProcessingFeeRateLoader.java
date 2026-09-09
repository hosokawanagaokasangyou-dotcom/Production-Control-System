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
 * 受注ﾌｧｲﾙシートから依頼No → 加工賃（AH・円/m）のマップを読む。
 *
 * <p>列位置は {@link JuchuSheetColumnLayout.Col} の index 固定。同一依頼No は後勝ち。
 * AH が空／非数値の行はマップに載せない。改行複数値は先頭の数値のみ採用する。
 */
public final class JuchuProcessingFeeRateLoader {

    public static final String SHEET_NAME = "受注ﾌｧｲﾙ";
    /** 見出し行（1-based）。工場既定に合わせる。 */
    public static final int DEFAULT_HEADER_ROW_ONE_BASED = 3;

    private static final DataFormatter FORMATTER = new DataFormatter(Locale.JAPAN);

    private JuchuProcessingFeeRateLoader() {}

    /**
     * @param workbookPath 受注 Excel（.xlsx / .xlsm）
     * @return 依頼No（trim）→ 加工賃単価。空マップもあり得る
     */
    public static Map<String, Double> loadRates(Path workbookPath) throws IOException {
        return loadRates(workbookPath, DEFAULT_HEADER_ROW_ONE_BASED);
    }

    public static Map<String, Double> loadRates(Path workbookPath, int headerRowOneBased)
            throws IOException {
        Objects.requireNonNull(workbookPath, "workbookPath");
        if (!Files.isRegularFile(workbookPath)) {
            throw new IOException("受注ファイルがありません: " + workbookPath);
        }
        try (InputStream in = Files.newInputStream(workbookPath);
                Workbook wb = WorkbookFactory.create(in)) {
            return loadRates(wb, headerRowOneBased);
        }
    }

    public static Map<String, Double> loadRates(Workbook wb, int headerRowOneBased) {
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
        int firstDataRow = headerRowOneBased; // 0-based: header is headerRowOneBased-1
        Map<String, Double> out = new LinkedHashMap<>();
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
            Double rate = parseFeeRate(cellText(row.getCell(feeCol), eval));
            if (rate == null) {
                continue;
            }
            out.put(irai, rate);
        }
        return Collections.unmodifiableMap(out);
    }

    /** 先頭行の数値を単価として返す。空・非数値は null。 */
    static Double parseFeeRate(String raw) {
        if (raw == null) {
            return null;
        }
        String s = raw.strip();
        if (s.isEmpty()) {
            return null;
        }
        int nl = s.indexOf('\n');
        if (nl >= 0) {
            s = s.substring(0, nl).strip();
        }
        int cr = s.indexOf('\r');
        if (cr >= 0) {
            s = s.substring(0, cr).strip();
        }
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
