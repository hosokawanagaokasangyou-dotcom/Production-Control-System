package jp.co.pm.ai.kouchin.verify;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * xlsx / xlsm を「値の二次元配列」として読むためのユーティリティ。
 * Python 版の python-calamine {@code to_python()} に相当し、
 * 数式セルはキャッシュされた計算結果を返す。
 */
public final class ExcelValues {

    private ExcelValues() {
    }

    /** 読み取り専用でブックを開く。呼び出し側で close すること。 */
    public static Workbook open(Path path) {
        try {
            return WorkbookFactory.create(path.toFile(), null, true);
        } catch (IOException | RuntimeException e) {
            throw new VerifyException("Excelファイルを開けません: " + path + " (" + e.getMessage() + ")", e);
        }
    }

    /** ブック内のシート名一覧。 */
    public static List<String> sheetNames(Workbook wb) {
        List<String> names = new ArrayList<>(wb.getNumberOfSheets());
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            names.add(wb.getSheetName(i));
        }
        return names;
    }

    /** シート名で値を読む。シートが無ければ空リスト。 */
    public static List<List<Object>> readSheet(Workbook wb, String sheetName) {
        Sheet sheet = wb.getSheet(sheetName);
        return sheet == null ? List.of() : readSheet(sheet);
    }

    /** 先頭シートの値を読む。 */
    public static List<List<Object>> readFirstSheet(Workbook wb) {
        if (wb.getNumberOfSheets() == 0) {
            return List.of();
        }
        return readSheet(wb.getSheetAt(0));
    }

    /** シートの値を行×列の二次元リストで返す。 */
    public static List<List<Object>> readSheet(Sheet sheet) {
        int lastRow = sheet.getLastRowNum();
        List<List<Object>> rows = new ArrayList<>(Math.max(0, lastRow + 1));
        for (int i = 0; i <= lastRow; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                rows.add(Collections.emptyList());
                continue;
            }
            int lastCell = row.getLastCellNum();
            List<Object> values = new ArrayList<>(Math.max(0, lastCell));
            for (int c = 0; c < lastCell; c++) {
                values.add(cellValue(row.getCell(c)));
            }
            rows.add(values);
        }
        return rows;
    }

    /** セル値（文字列 / Double / Boolean / null）。数式はキャッシュ値を使う。 */
    public static Object cellValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        CellType type = cell.getCellType();
        if (type == CellType.FORMULA) {
            try {
                type = cell.getCachedFormulaResultType();
            } catch (RuntimeException e) {
                return null;
            }
        }
        try {
            return switch (type) {
                case STRING -> cell.getStringCellValue();
                case NUMERIC -> cell.getNumericCellValue();
                case BOOLEAN -> cell.getBooleanCellValue();
                default -> null;
            };
        } catch (RuntimeException e) {
            return null;
        }
    }

    /** 0始まりの安全なセル取得。範囲外は null。 */
    public static Object at(List<List<Object>> rows, int rowIndex, int colIndex) {
        if (rowIndex < 0 || rowIndex >= rows.size()) {
            return null;
        }
        List<Object> row = rows.get(rowIndex);
        if (colIndex < 0 || colIndex >= row.size()) {
            return null;
        }
        return row.get(colIndex);
    }

    /** 1始まり（Excel の行番号・列番号）の安全なセル取得。 */
    public static Object cell1(List<List<Object>> rows, int row1, int col1) {
        return at(rows, row1 - 1, col1 - 1);
    }

    /** 1始まりで数値を取得。数値でなければ 0。 */
    public static double num1(List<List<Object>> rows, int row1, int col1) {
        Double v = Norm.number(cell1(rows, row1, col1));
        return v == null ? 0.0 : v;
    }
}
