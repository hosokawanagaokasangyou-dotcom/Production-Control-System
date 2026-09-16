package jp.co.pm.ai.kouchin.trend;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.kouchin.verify.YearMonthKey;

class TrendExcelWriterTest {

    @TempDir Path tmp;

    @Test
    @DisplayName("シートはサマリ/工程比較/負荷シフト/付録3種。工程比較のみチャート")
    void writeCreatesExpectedSheetsAndCompareHasCharts() throws Exception {
        Path out = tmp.resolve("月トレンド_test.xlsx");
        try (XSSFWorkbook built = TrendExcelWriter.buildWorkbook(minimalModel())) {
            try (var os = Files.newOutputStream(out)) {
                built.write(os);
            }
        }
        try (InputStream in = Files.newInputStream(out);
                XSSFWorkbook wb = new XSSFWorkbook(in)) {
            assertEquals(
                    List.of("サマリ", "工程比較", "負荷シフト", "付録_国分", "付録_湖南", "付録_シート区分"),
                    sheetNames(wb));
            XSSFSheet compare = wb.getSheet("工程比較");
            XSSFDrawing compareDraw = compare.getDrawingPatriarch();
            assertNotNull(compareDraw);
            assertFalse(compareDraw.getCharts().isEmpty());
            XSSFSheet appK = wb.getSheet("付録_国分");
            XSSFDrawing appDraw = appK.getDrawingPatriarch();
            assertTrue(appDraw == null || appDraw.getCharts().isEmpty());
            XSSFChart chart0 = compareDraw.getCharts().get(0);
            assertEquals(2, chart0.getCTChart().getPlotArea().getLineChartArray(0).sizeOfSerArray());
            assertFalse(chart0.getTitleOverlay());
            assertEquals("区分", cellString(compare, 6, 0));
            List<Integer> chartRows = new ArrayList<>();
            var ctDrawing = compareDraw.getCTDrawing();
            if (ctDrawing != null) {
                for (var anchor : ctDrawing.getTwoCellAnchorList()) {
                    if (anchor.getFrom() != null) {
                        chartRows.add(anchor.getFrom().getRow());
                    }
                }
            }
            for (int i = 0; i < chartRows.size() - 1; i++) {
                assertTrue(
                        chartRows.get(i + 1) - chartRows.get(i) >= 16,
                        "Chart overlap: row " + chartRows.get(i) + " vs " + chartRows.get(i + 1));
            }
            List<String> cmpTexts = colA(compare, 40);
            assertTrue(cmpTexts.stream().anyMatch(t -> t.contains("km")));
            assertTrue(cmpTexts.stream().anyMatch(t -> t.contains("千円")));
            String flat = String.join(" ", colA(wb.getSheet("サマリ"), 40));
            assertTrue(flat.contains("工程比較"));
            assertTrue(flat.contains("スライス") || flat.contains("振替"));
        }
    }

    private static TrendModel minimalModel() {
        YearMonthKey jul = new YearMonthKey(2026, 7);
        YearMonthKey aug = new YearMonthKey(2026, 8);
        TrendMonthData k7 =
                new TrendMonthData(
                        jul,
                        Path.of("k7.xlsx"),
                        Map.of(
                                "スライス1", new TrendMetric(100, 10),
                                "スリット", new TrendMetric(50, 5)),
                        Map.of("東レT", new TrendMetric(150, 15)),
                        List.of());
        TrendMonthData k8 =
                new TrendMonthData(
                        aug,
                        Path.of("k8.xlsx"),
                        Map.of(
                                "スライス1", new TrendMetric(180, 18),
                                "スリット", new TrendMetric(55, 5)),
                        Map.of("東レT", new TrendMetric(235, 23)),
                        List.of());
        TrendMonthData n7 =
                new TrendMonthData(
                        jul,
                        Path.of("n7.xlsx"),
                        Map.of(
                                "スライス", new TrendMetric(200, 20),
                                "スリット", new TrendMetric(40, 4)),
                        Map.of("東レT.V.C", new TrendMetric(240, 24)),
                        List.of());
        TrendMonthData n8 =
                new TrendMonthData(
                        aug,
                        Path.of("n8.xlsx"),
                        Map.of(
                                "スライス", new TrendMetric(120, 12),
                                "スリット", new TrendMetric(42, 4)),
                        Map.of("東レT.V.C", new TrendMetric(162, 16)),
                        List.of());
        return TrendModel.assemble(
                List.of(jul, aug), List.of(k7, k8), List.of(n7, n8), List.of("テスト警告"));
    }

    private static List<String> sheetNames(XSSFWorkbook wb) {
        List<String> names = new ArrayList<>();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            names.add(wb.getSheetName(i));
        }
        return names;
    }

    private static String cellString(XSSFSheet sheet, int row0, int col0) {
        Row row = sheet.getRow(row0);
        if (row == null) {
            return null;
        }
        Cell cell = row.getCell(col0);
        return cell == null ? null : cell.getStringCellValue();
    }

    private static List<String> colA(XSSFSheet sheet, int maxRow) {
        List<String> texts = new ArrayList<>();
        for (int i = 0; i < maxRow; i++) {
            Row row = sheet.getRow(i);
            if (row == null || row.getCell(0) == null) {
                continue;
            }
            Cell cell = row.getCell(0);
            if (cell.getCellType() == org.apache.poi.ss.usermodel.CellType.STRING
                    && !cell.getStringCellValue().isBlank()) {
                texts.add(cell.getStringCellValue());
            }
        }
        return texts;
    }
}
