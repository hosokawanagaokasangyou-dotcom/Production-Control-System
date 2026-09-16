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

class TrendKonanReaderTest {

    @TempDir Path tmp;

    @Test
    @DisplayName("東レ3シートの種類別賃・量を合算する")
    void readKonanKindsSumAcrossSheets() throws Exception {
        Path path = tmp.resolve("8月度加工賃試算.xlsx");
        makeKonan(path);
        TrendMonthData agg = TrendKonanReader.read(path);
        assertEquals(new YearMonthKey(2026, 8), agg.ym());
        assertEquals(150.0, agg.kinds().get("スリット").wage(), 1e-9);
        assertEquals(15.0, agg.kinds().get("スリット").qty(), 1e-9);
        assertEquals(100.0, agg.sheets().get("東レT.V.C").wage(), 1e-9);
        assertEquals(50.0, agg.sheets().get("東レY.S").wage(), 1e-9);
        assertEquals(10.0, agg.sheets().get("東レT.V.C").qty(), 1e-9);
    }

    private static void makeKonan(Path path) throws Exception {
        Files.createDirectories(path.getParent());
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            addToraySheet(wb, "東レT.V.C", "2026年8月度", 100, 10, 10);
            addToraySheet(wb, "東レY.S", "2026年8月度", 50, 5, 5);
            addToraySheet(wb, "東レW.E..", "2026年8月度", 0, 0, 0);
            try (var out = Files.newOutputStream(path)) {
                wb.write(out);
            }
        }
    }

    private static void addToraySheet(
            XSSFWorkbook wb, String title, String ymLabel, double slitWage, double slitQty, double inbound) {
        XSSFSheet ws = wb.createSheet(title);
        cell(ws, 0, 0, ymLabel);
        cell(ws, 7, 0, "加工依頼No.");
        cell(ws, 7, 2, "契約No.");
        cell(ws, 7, 3, "入庫数量");
        cell(ws, 7, 4, "加工内容");
        cell(ws, 7, 21, "加工種類別 加工賃");
        cell(ws, 7, 37, "加工種類別 加工量");
        cell(ws, 8, 5, "合計");
        cell(ws, 8, 6, "スリット");
        cell(ws, 8, 21, "合計");
        cell(ws, 8, 22, "スリット");
        cell(ws, 8, 37, "合計");
        cell(ws, 8, 38, "スリット");
        cell(ws, 9, 6, 10);
        cell(ws, 10, 21, 0);
        cell(ws, 11, 0, "C8-1");
        cell(ws, 11, 2, "191352R");
        cell(ws, 11, 3, inbound);
        cell(ws, 11, 4, "スリット");
        cell(ws, 11, 21, slitWage);
        cell(ws, 11, 22, slitWage);
        cell(ws, 11, 37, slitQty);
        cell(ws, 11, 38, slitQty);
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
