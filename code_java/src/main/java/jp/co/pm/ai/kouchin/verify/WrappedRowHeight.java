package jp.co.pm.ai.kouchin.verify;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * 折り返しコメントが見切れる行だけ高さを伸ばす。短い行はそのままなので、シート内の行高は揃わなくてよい。
 */
public final class WrappedRowHeight {

    private static final float MAX_POINTS = 409f;

    private WrappedRowHeight() {}

    public static void fit(XSSFSheet sheet) {
        if (sheet == null) {
            return;
        }
        for (Row row : sheet) {
            float needed = 0f;
            for (Cell cell : row) {
                if (cell == null || cell.getCellType() != CellType.STRING) {
                    continue;
                }
                if (!(cell.getCellStyle() instanceof XSSFCellStyle style) || !style.getWrapText()) {
                    continue;
                }
                String text = cell.getStringCellValue();
                if (text == null || text.isBlank()) {
                    continue;
                }
                int width = widthUnits(sheet, row.getRowNum(), cell.getColumnIndex());
                if (width <= 0) {
                    continue;
                }
                int lines = lines(text, width);
                if (lines <= 1) {
                    continue;
                }
                float fontPt = style.getFont().getFontHeightInPoints();
                float linePt = Math.max(13f, fontPt + 4f);
                needed = Math.max(needed, lines * linePt + 3f);
            }
            if (needed > row.getHeightInPoints() + 0.5f) {
                row.setHeightInPoints(Math.min(MAX_POINTS, needed));
            }
        }
    }

    /** 列幅（1/256文字）に対して、全角を2として必要な行数を返す。 */
    static int lines(String text, int columnWidth256) {
        double capacity = Math.max(4.0, columnWidth256 / 256.0 - 0.5);
        int lines = 0;
        for (String part : text.split("\n", -1)) {
            double used = 0;
            int rowLines = 1;
            for (int i = 0; i < part.length(); i++) {
                double width = part.charAt(i) > 0xFF ? 2.0 : 1.0;
                if (used + width > capacity && used > 0) {
                    rowLines++;
                    used = 0;
                }
                used += width;
            }
            lines += Math.max(1, rowLines);
        }
        return Math.max(1, lines);
    }

    private static int widthUnits(XSSFSheet sheet, int rowIndex, int col) {
        int start = col;
        int end = col;
        for (CellRangeAddress region : sheet.getMergedRegions()) {
            if (region.isInRange(rowIndex, col)) {
                if (region.getFirstColumn() != col) {
                    return 0;
                }
                start = region.getFirstColumn();
                end = region.getLastColumn();
                break;
            }
        }
        int sum = 0;
        for (int c = start; c <= end; c++) {
            sum += sheet.getColumnWidth(c);
        }
        return sum;
    }
}
