package jp.co.pm.ai.desktop.reconciliation;

import java.util.Locale;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * 依頼書原本シートの加工賃（円）。
 * 製品数量（AE10–AE12 の m）× 加工内容の単価合計（P13–P17 の円/m）。
 */
public final class RequestFormOriginalFee {

    private static final DataFormatter FORMATTER = new DataFormatter(Locale.JAPAN);
    /** 加工内容の単価列（円/m）。Excel P 列。 */
    private static final int PRICE_COLUMN = RequestFormOriginalCellLayout.columnLetterToIndex("P");

    /** @param meters 数量合計（m） @param yenPerMeter 工程単価の合計（円/m） @param amountYen 加工賃（円） */
    public record Result(double meters, double yenPerMeter, double amountYen) {}

    private RequestFormOriginalFee() {}

    /** 数量も単価も無いときは empty。 */
    public static Result fromSheet(Sheet sheet) {
        if (sheet == null) {
            return null;
        }
        double meters = 0.0;
        int qtyCol = RequestFormOriginalCellLayout.ProductColumn.QTY.columnIndex();
        for (int rowIndex : RequestFormOriginalCellLayout.PRODUCT_ROW_INDICES) {
            meters += cellNumber(sheet, rowIndex, qtyCol);
        }
        double rate = 0.0;
        for (int rowIndex : RequestFormOriginalCellLayout.PROCESS_STEP_ROW_INDICES) {
            rate += cellNumber(sheet, rowIndex, PRICE_COLUMN);
        }
        if (meters <= 0.0 || rate <= 0.0) {
            return null;
        }
        return new Result(meters, rate, meters * rate);
    }

    private static double cellNumber(Sheet sheet, int rowIndex, int columnIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return 0.0;
        }
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            return 0.0;
        }
        CellType type = cell.getCellType();
        if (type == CellType.FORMULA) {
            try {
                type = cell.getCachedFormulaResultType();
            } catch (RuntimeException ex) {
                return 0.0;
            }
        }
        if (type == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        }
        String text;
        try {
            text = FORMATTER.formatCellValue(cell);
        } catch (RuntimeException ex) {
            return 0.0;
        }
        if (text == null) {
            return 0.0;
        }
        text = text.replace(",", "").replace("，", "").strip();
        if (text.isEmpty()) {
            return 0.0;
        }
        try {
            return Double.parseDouble(text);
        } catch (NumberFormatException ex) {
            return 0.0;
        }
    }
}
