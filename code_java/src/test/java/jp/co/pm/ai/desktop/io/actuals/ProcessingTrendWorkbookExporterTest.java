package jp.co.pm.ai.desktop.io.actuals;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.DayPoint;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.Filter;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.PlanSource;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.Result;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendWorkbookExporter.ProcessingTrendExportRequest;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendWorkbookExporter.ProcessingTrendExportResult;

class ProcessingTrendWorkbookExporterTest {

    @Test
    void suggestFileName_containsPeriodAndSanitizedConditions() {
        Filter f =
                new Filter(
                        LocalDate.of(2026, 9, 1),
                        LocalDate.of(2026, 9, 30),
                        PlanSource.ALADDIN,
                        "SEC機/湖南",
                        "スリット*test");
        LocalDateTime at = LocalDateTime.of(2026, 9, 7, 10, 30, 0);

        String fn = ProcessingTrendWorkbookExporter.suggestFileName(f, false, at);
        assertTrue(fn.startsWith("加工トレンド_20260901-20260930_日別_"));
        assertTrue(fn.contains("SEC機_湖南"));
        assertTrue(fn.contains("スリット_test"));
        assertTrue(fn.endsWith(".xlsx"));
        assertTrue(!fn.contains("/"));
        assertTrue(!fn.contains("*"));

        String fnMonth = ProcessingTrendWorkbookExporter.suggestFileName(f, true, at);
        assertTrue(fnMonth.startsWith("加工トレンド_20260901-20260930_月別_"));
    }

    @Test
    void buildWorkbook_createsSummaryAndDailySheetsWithCorrectData(@TempDir Path tempDir) throws Exception {
        LocalDate start = LocalDate.of(2026, 9, 1);
        LocalDate end = LocalDate.of(2026, 9, 3);
        LocalDate today = LocalDate.of(2026, 9, 2);
        Filter f = new Filter(start, end, PlanSource.ALADDIN, "W9-1", "スリット");

        List<DayPoint> days = new ArrayList<>();
        days.add(new DayPoint(LocalDate.of(2026, 9, 1), 100.0, 120.0, 100.0, 120.0, 100.0, false));
        days.add(new DayPoint(LocalDate.of(2026, 9, 2), 150.0, 80.0, 250.0, 200.0, 250.0, true));
        days.add(new DayPoint(LocalDate.of(2026, 9, 3), 0.0, 200.0, 250.0, 400.0, 450.0, true));

        Result result =
                new Result(
                        days, 250.0, 400.0, 100.0, 120.0, 200.0, 450.0,
                        today, 5, 4, start, end, List.of("テスト注意"));

        ProcessingTrendAggregator.MonthlyResult mr =
                ProcessingTrendAggregator.rollUpMonthly(result, f, today);

        LocalDateTime now = LocalDateTime.of(2026, 9, 7, 12, 0, 0);
        ProcessingTrendExportRequest req =
                new ProcessingTrendExportRequest(
                        result,
                        mr,
                        f,
                        now,
                        "actual-detail-newest.xlsx",
                        "アラジン加工計画",
                        "task-input-newest.xlsx",
                        "結果_配台表.json",
                        "行 実績 10 / アラジン 5",
                        List.of("テスト注意文"));

        try (XSSFWorkbook wb = ProcessingTrendWorkbookExporter.buildWorkbook(req)) {
            assertTrue(wb.getNumberOfSheets() >= 4, "サマリ・日別・機械別・機械工程別があること");
            Sheet sSummary = wb.getSheet(ProcessingTrendWorkbookExporter.SHEET_SUMMARY);
            Sheet sDaily = wb.getSheet(ProcessingTrendWorkbookExporter.SHEET_DAILY);
            assertNotNull(sSummary, "サマリシートが存在すること");
            assertNotNull(sDaily, "日別明細シートが存在すること");

            // サマリシートのタイトルと条件
            assertEquals("加工トレンド集計サマリ", sSummary.getRow(0).getCell(0).getStringCellValue());
            assertTrue(sSummary.getRow(3).getCell(1).getStringCellValue().contains("2026/09/01 〜 2026/09/03"));

            // 日別明細シート
            assertEquals("日付", sDaily.getRow(0).getCell(0).getStringCellValue());
            assertEquals("曜日", sDaily.getRow(0).getCell(1).getStringCellValue());
            assertEquals("実績 (m)", sDaily.getRow(0).getCell(2).getStringCellValue());
            assertEquals("30日移動平均 (m)", sDaily.getRow(0).getCell(3).getStringCellValue());
            assertEquals("予定 (m)", sDaily.getRow(0).getCell(4).getStringCellValue());
            assertEquals("実績累計 (m)", sDaily.getRow(0).getCell(5).getStringCellValue());

            // 整数書式 (#,##0" m")
            String fmt =
                    wb.createDataFormat()
                            .getFormat(sDaily.getRow(1).getCell(2).getCellStyle().getDataFormat());
            assertTrue(fmt.contains("#,##0"), "整数書式であること: " + fmt);
            assertTrue(!fmt.contains(".0"), "小数書式ではないこと: " + fmt);

            // トレンドチャートが日別シートにある
            assertNotNull(((org.apache.poi.xssf.usermodel.XSSFSheet) sDaily).getDrawingPatriarch());
            assertFalse(
                    ((org.apache.poi.xssf.usermodel.XSSFSheet) sDaily)
                            .getDrawingPatriarch()
                            .getCharts()
                            .isEmpty());

            assertNotNull(wb.getSheet(ProcessingTrendWorkbookExporter.SHEET_BY_MACHINE));
            assertNotNull(wb.getSheet(ProcessingTrendWorkbookExporter.SHEET_BY_MACHINE_PROCESS));


            // データ行3日分 + 合計行1行
            // 行0: ヘッダー, 行1: 9/1, 行2: 9/2, 行3: 9/3, 行4: 合計
            assertNotNull(sDaily.getRow(4), "合計行が存在すること");
            assertEquals("合計", sDaily.getRow(4).getCell(0).getStringCellValue());
            assertEquals(250.0, sDaily.getRow(4).getCell(2).getNumericCellValue(), 1e-9);
            assertEquals(400.0, sDaily.getRow(4).getCell(4).getNumericCellValue(), 1e-9);
        }

        Path target = tempDir.resolve("export_test.xlsx");
        ProcessingTrendExportResult expResult =
                ProcessingTrendWorkbookExporter.writeTo(target, req, Map.of());
        assertTrue(Files.exists(target));
        assertTrue(Files.size(target) > 0);
        assertEquals(3, expResult.dayRows());
    }

    @Test
    void buildWorkbook_withComparison_outputsDailyReportAndDetailColumns(@TempDir Path tempDir) throws Exception {
        LocalDate start = LocalDate.of(2026, 9, 1);
        LocalDate end = LocalDate.of(2026, 9, 2);
        LocalDate today = LocalDate.of(2026, 9, 2);
        Filter f =
                new Filter(
                        start,
                        end,
                        ProcessingTrendAggregator.ActualSource.DAILY_REPORT,
                        PlanSource.ALADDIN,
                        "W9-1",
                        "スリット");

        List<DayPoint> days = new ArrayList<>();
        // 日報 120, 明細 100 -> 差 20
        days.add(
                new DayPoint(
                        LocalDate.of(2026, 9, 1),
                        120.0,
                        100.0,
                        120.0,
                        100.0,
                        120.0,
                        false,
                        100.0,
                        100.0));
        // 日報 150, 明細 130 -> 差 20
        days.add(
                new DayPoint(
                        LocalDate.of(2026, 9, 2),
                        150.0,
                        80.0,
                        270.0,
                        180.0,
                        270.0,
                        true,
                        130.0,
                        230.0));

        Result result =
                new Result(
                        days, 270.0, 180.0, 120.0, 100.0, 0.0, 270.0,
                        today, 2, 2, start, end, List.of(), 230.0, "実績明細");

        ProcessingTrendAggregator.MonthlyResult mr =
                ProcessingTrendAggregator.rollUpMonthly(result, f, today);

        ProcessingTrendExportRequest req =
                new ProcessingTrendExportRequest(
                        result,
                        mr,
                        f,
                        LocalDateTime.of(2026, 9, 7, 12, 0, 0),
                        "actual-detail-newest.xlsx",
                        "アラジン加工計画",
                        "task-input-newest.xlsx",
                        "結果_配台表.json",
                        "",
                        List.of(),
                        "加工日報発行問合せ_20260904.csv");

        try (XSSFWorkbook wb = ProcessingTrendWorkbookExporter.buildWorkbook(req)) {
            Sheet sDaily = wb.getSheet(ProcessingTrendWorkbookExporter.SHEET_DAILY);
            assertNotNull(sDaily);

            // 日報実績 (m), 30日移動平均 (m), 実績明細 (m), 差異 (m), 予定 (m)
            assertEquals("日付", sDaily.getRow(0).getCell(0).getStringCellValue());
            assertEquals("曜日", sDaily.getRow(0).getCell(1).getStringCellValue());
            assertEquals("日報実績 (m)", sDaily.getRow(0).getCell(2).getStringCellValue());
            assertEquals("30日移動平均 (m)", sDaily.getRow(0).getCell(3).getStringCellValue());
            assertEquals("実績明細 (m)", sDaily.getRow(0).getCell(4).getStringCellValue());
            assertEquals("差異 (m)", sDaily.getRow(0).getCell(5).getStringCellValue());
            assertEquals("予定 (m)", sDaily.getRow(0).getCell(6).getStringCellValue());
            assertEquals("日報累計 (m)", sDaily.getRow(0).getCell(7).getStringCellValue());

            // 9/1 データ行
            assertEquals(120.0, sDaily.getRow(1).getCell(2).getNumericCellValue(), 1e-9);
            assertEquals(100.0, sDaily.getRow(1).getCell(4).getNumericCellValue(), 1e-9);
            assertEquals(20.0, sDaily.getRow(1).getCell(5).getNumericCellValue(), 1e-9);

            // 合計行
            // 行0: ヘッダー, 行1: 9/1, 行2: 9/2, 行3: 合計
            assertEquals("合計", sDaily.getRow(3).getCell(0).getStringCellValue());
            assertEquals(270.0, sDaily.getRow(3).getCell(2).getNumericCellValue(), 1e-9);
            assertEquals(230.0, sDaily.getRow(3).getCell(4).getNumericCellValue(), 1e-9);
            assertEquals(40.0, sDaily.getRow(3).getCell(5).getNumericCellValue(), 1e-9);
        }
    }
}
