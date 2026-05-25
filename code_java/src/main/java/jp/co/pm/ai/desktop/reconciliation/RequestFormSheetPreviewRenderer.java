package jp.co.pm.ai.desktop.reconciliation;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * 依頼書 Excel シートを Apache POI で読み、Apache PDFBox で PDF プレビューを生成する。
 * xlwings / Excel COM は使用しない。
 */
final class RequestFormSheetPreviewRenderer {

    /** Excel 1-based プレビュー範囲（依頼書フォーム全体）。 */
    static final String PREVIEW_RANGE_SPEC = "A1:AO29";

    /** 描画方式の識別子（変更時はプレビューキャッシュを無効化する）。 */
    static final String PREVIEW_RENDERER_SPEC = RequestFormSheetPreviewPdfRenderer.RENDERER_SPEC;

    /** POI 0-based: 行 1～29 → 0～28。 */
    private static final int PREVIEW_FIRST_ROW = 0;
    private static final int PREVIEW_LAST_ROW = 28;

    /** POI 0-based: 列 A～AO → 0～40。 */
    private static final int PREVIEW_FIRST_COL = 0;
    private static final int PREVIEW_LAST_COL = 40;

    private static final double DEFAULT_COL_WIDTH_PX = 64.0;
    private static final double DEFAULT_ROW_HEIGHT_PX = 24.0;
    private static final double MIN_ROW_HEIGHT_PX = 12.0;

    private RequestFormSheetPreviewRenderer() {}

    record PreviewData(
            int firstRow,
            int firstCol,
            int rowCount,
            int colCount,
            String[][] texts,
            RequestFormPreviewCellStyle[][] styles,
            List<RequestFormPreviewTextRun>[][] richRuns,
            HorizontalAlignment[][] hAligns,
            boolean[][] skip,
            int[][] colSpans,
            int[][] rowSpans,
            double[] colWidthsPx,
            double[] rowHeightsPx,
            List<RequestFormSheetShapeOverlay.OverlayShape> overlayShapes) {}

    static void generatePreviewPdf(File excelFile, String sheetName, File outputFile) throws Exception {
        RequestFormSheetPreviewPdfRenderer.generatePreviewPdf(excelFile, sheetName, outputFile);
    }

    @SuppressWarnings("unchecked")
    static PreviewData loadPreviewData(File excelFile, String sheetName) throws IOException {
        try (FileInputStream fis = new FileInputStream(excelFile);
                Workbook wb = WorkbookFactory.create(fis)) {
            Sheet sheet = wb.getSheet(sheetName);
            if (sheet == null) {
                throw new IOException("シートが見つかりません: " + sheetName);
            }

            int firstRow = PREVIEW_FIRST_ROW;
            int lastRow = PREVIEW_LAST_ROW;
            int firstCol = PREVIEW_FIRST_COL;
            int lastCol = PREVIEW_LAST_COL;

            int rowCount = lastRow - firstRow + 1;
            int colCount = lastCol - firstCol + 1;

            String[][] texts = new String[rowCount][colCount];
            RequestFormPreviewCellStyle[][] styles = new RequestFormPreviewCellStyle[rowCount][colCount];
            List<RequestFormPreviewTextRun>[][] richRuns = new List[rowCount][colCount];
            HorizontalAlignment[][] hAligns = new HorizontalAlignment[rowCount][colCount];
            boolean[][] skip = new boolean[rowCount][colCount];
            int[][] colSpans = new int[rowCount][colCount];
            int[][] rowSpans = new int[rowCount][colCount];
            double[] colWidthsPx = new double[colCount];
            double[] rowHeightsPx = new double[rowCount];
            double defaultRowPx = defaultRowHeightPx(sheet);

            for (int r = 0; r < rowCount; r++) {
                rowHeightsPx[r] = defaultRowPx;
            }
            for (int c = 0; c < colCount; c++) {
                colWidthsPx[c] = columnWidthPx(sheet, firstCol + c);
            }

            for (int r = 0; r < rowCount; r++) {
                colSpans[r] = new int[colCount];
                rowSpans[r] = new int[colCount];
                for (int c = 0; c < colCount; c++) {
                    colSpans[r][c] = 1;
                    rowSpans[r][c] = 1;
                    hAligns[r][c] = HorizontalAlignment.LEFT;
                    styles[r][c] = RequestFormPreviewCellStyle.defaults();
                }
            }

            Map<String, CellRangeAddress> mergeByDisplayAnchor = new HashMap<>();
            for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                CellRangeAddress region = sheet.getMergedRegion(i);
                int clipR1 = Math.max(region.getFirstRow(), firstRow);
                int clipC1 = Math.max(region.getFirstColumn(), firstCol);
                int clipR2 = Math.min(region.getLastRow(), lastRow);
                int clipC2 = Math.min(region.getLastColumn(), lastCol);
                if (clipR1 > clipR2 || clipC1 > clipC2) {
                    continue;
                }
                int startR = clipR1 - firstRow;
                int startC = clipC1 - firstCol;
                int endR = clipR2 - firstRow;
                int endC = clipC2 - firstCol;
                colSpans[startR][startC] = endC - startC + 1;
                rowSpans[startR][startC] = endR - startR + 1;
                mergeByDisplayAnchor.put(startR + "," + startC, region);
                for (int r = startR; r <= endR; r++) {
                    for (int c = startC; c <= endC; c++) {
                        if (r != startR || c != startC) {
                            skip[r][c] = true;
                        }
                    }
                }
            }

            DataFormatter formatter = new DataFormatter(Locale.JAPAN);
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            Map<Integer, Font> fontCache = new HashMap<>();

            for (int r = firstRow; r <= lastRow; r++) {
                Row row = sheet.getRow(r);
                int gridR = r - firstRow;
                rowHeightsPx[gridR] = rowHeightPx(sheet, row);
                for (int c = firstCol; c <= lastCol; c++) {
                    int gridC = c - firstCol;
                    CellRangeAddress mergeRegion = mergeByDisplayAnchor.get(gridR + "," + gridC);
                    Cell cell;
                    if (mergeRegion != null) {
                        Row anchorRow = sheet.getRow(mergeRegion.getFirstRow());
                        cell =
                                anchorRow != null
                                        ? anchorRow.getCell(mergeRegion.getFirstColumn())
                                        : null;
                    } else {
                        cell = row != null ? row.getCell(c) : null;
                    }
                    texts[gridR][gridC] =
                            cell != null ? formatter.formatCellValue(cell, evaluator) : "";
                    CellStyle cellStyle = cell != null ? cell.getCellStyle() : null;
                    styles[gridR][gridC] =
                            RequestFormPreviewStyleHelper.cellStyle(wb, cellStyle, fontCache);
                    List<RequestFormPreviewTextRun> runs =
                            RequestFormPreviewStyleHelper.richTextRuns(wb, cell, cellStyle, fontCache);
                    if (runs.size() > 1) {
                        richRuns[gridR][gridC] = runs;
                    }
                    hAligns[gridR][gridC] =
                            cellStyle != null ? cellStyle.getAlignment() : HorizontalAlignment.LEFT;
                }
            }

            for (int r = 0; r < rowCount; r++) {
                boolean hasContent = false;
                for (int c = 0; c < colCount; c++) {
                    if (!skip[r][c] && texts[r][c] != null && !texts[r][c].isBlank()) {
                        hasContent = true;
                        break;
                    }
                }
                if (hasContent && rowHeightsPx[r] < MIN_ROW_HEIGHT_PX) {
                    rowHeightsPx[r] = MIN_ROW_HEIGHT_PX;
                }
            }

            List<RequestFormSheetShapeOverlay.OverlayShape> overlayShapes = List.of();
            if (sheet instanceof XSSFSheet xssfSheet) {
                overlayShapes =
                        RequestFormSheetShapeOverlay.loadShapes(
                                xssfSheet,
                                firstRow,
                                lastRow,
                                firstCol,
                                lastCol,
                                colWidthsPx,
                                rowHeightsPx);
            }

            return new PreviewData(
                    firstRow,
                    firstCol,
                    rowCount,
                    colCount,
                    texts,
                    styles,
                    richRuns,
                    hAligns,
                    skip,
                    colSpans,
                    rowSpans,
                    colWidthsPx,
                    rowHeightsPx,
                    overlayShapes);
        }
    }

    private static double columnWidthPx(Sheet sheet, int col) {
        if (sheet instanceof XSSFSheet xssfSheet) {
            float px = xssfSheet.getColumnWidthInPixels(col);
            if (px > 0) {
                return Math.max(1.0, px);
            }
        }
        int width = sheet.getColumnWidth(col);
        if (width <= 0) {
            return defaultColumnWidthPx(sheet);
        }
        return Math.max(1.0, width * 7.0 / 256.0 + 5.0);
    }

    private static double defaultColumnWidthPx(Sheet sheet) {
        if (sheet instanceof XSSFSheet xssfSheet) {
            int defaultChars = xssfSheet.getDefaultColumnWidth();
            if (defaultChars > 0) {
                return Math.max(1.0, defaultChars * 7.0 + 5.0);
            }
        }
        return DEFAULT_COL_WIDTH_PX;
    }

    private static double defaultRowHeightPx(Sheet sheet) {
        float points = sheet.getDefaultRowHeightInPoints();
        if (points > 0) {
            return points * 96.0 / 72.0;
        }
        return DEFAULT_ROW_HEIGHT_PX;
    }

    private static double rowHeightPx(Sheet sheet, Row row) {
        if (row != null && row.getZeroHeight()) {
            return MIN_ROW_HEIGHT_PX;
        }
        if (row != null && row.getHeightInPoints() > 0) {
            return Math.max(MIN_ROW_HEIGHT_PX, row.getHeightInPoints() * 96.0 / 72.0);
        }
        return defaultRowHeightPx(sheet);
    }
}
