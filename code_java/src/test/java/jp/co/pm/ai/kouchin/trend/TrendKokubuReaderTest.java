package jp.co.pm.ai.kouchin.trend;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.nio.file.Files;
import java.nio.file.Path;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.kouchin.verify.YearMonthKey;

class TrendKokubuReaderTest {

    @TempDir Path tmp;

    @Test
    @DisplayName("東レまとめの工程別賃・量と東レTのシート合計を読む")
    void readKokubuKindsAndSheets() throws Exception {
        Path path = tmp.resolve("後加工工賃明細（2026年8月度)V01.xlsx");
        makeKokubu(path);
        TrendMonthData agg = TrendKokubuReader.read(path);
        assertEquals(new YearMonthKey(2026, 8), agg.ym());
        assertEquals(1000.0, agg.kinds().get("スリット").wage(), 1e-9);
        assertEquals(10.0, agg.kinds().get("スリット").qty(), 1e-9);
        assertEquals(500.0, agg.kinds().get("カット").wage(), 1e-9);
        assertEquals(5.0, agg.kinds().get("カット").qty(), 1e-9);
        assertEquals(1500.0, agg.sheets().get("東レT").wage(), 1e-9);
        assertEquals(15.0, agg.sheets().get("東レT").qty(), 1e-9);
    }

    private static void makeKokubu(Path path) throws Exception {
        Files.createDirectories(path.getParent());
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet ws = wb.createSheet("東レまとめ");
            cell(ws, 2, 0, "加工依頼No.");
            cell(ws, 2, 2, "契約No.");
            cell(ws, 2, 4, "入庫数量");
            cell(ws, 2, 26, "加工種類別 加工賃");
            cell(ws, 2, 45, "加工種類別 入庫量");
            cell(ws, 3, 6, "合計");
            cell(ws, 3, 7, "スリット");
            cell(ws, 3, 8, "カット");
            cell(ws, 3, 26, "合計");
            cell(ws, 3, 27, "スリット");
            cell(ws, 3, 28, "カット");
            cell(ws, 3, 45, "スリット");
            cell(ws, 3, 46, "カット");
            cell(ws, 5, 0, "T8");
            cell(ws, 5, 1, 1);
            cell(ws, 5, 2, "191352R");
            cell(ws, 5, 4, 15);
            cell(ws, 5, 26, 1500);
            cell(ws, 5, 27, 1000);
            cell(ws, 5, 28, 500);
            cell(ws, 5, 45, 10);
            cell(ws, 5, 46, 5);
            cell(ws, 6, 2, "加工賃合計");
            cell(ws, 6, 26, 9999);
            XSSFSheet t = wb.createSheet("東レT");
            cell(t, 5, 0, "T8");
            cell(t, 5, 1, 1);
            cell(t, 5, 2, "191352R");
            cell(t, 5, 4, 15);
            cell(t, 5, 26, 1500);
            wb.createSheet("東レV.C");
            wb.createSheet("東レY");
            wb.createSheet("東レW.E");
            try (var out = Files.newOutputStream(path)) {
                wb.write(out);
            }
        }
    }

    private static void cell(XSSFSheet ws, int row0, int col0, String v) {
        row(ws, row0).createCell(col0).setCellValue(v);
    }

    private static void cell(XSSFSheet ws, int row0, int col0, double v) {
        row(ws, row0).createCell(col0).setCellValue(v);
    }

    private static Row row(XSSFSheet ws, int row0) {
        Row row = ws.getRow(row0);
        return row != null ? row : ws.createRow(row0);
    }
}
