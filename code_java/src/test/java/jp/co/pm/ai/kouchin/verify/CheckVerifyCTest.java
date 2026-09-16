package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class CheckVerifyCTest {

    @TempDir Path tmp;

    @Test
    void matchExplainableNeedCheckUnreadable() throws Exception {
        Map<String, Double> monthly = new LinkedHashMap<>();
        monthly.put(CheckVerifyC.ITEM_JISSEKI1, 100.0);
        monthly.put(CheckVerifyC.ITEM_SHUKEI, 100.0);
        monthly.put(CheckVerifyC.ITEM_JISSEKI2, 80.0);
        monthly.put(CheckVerifyC.ITEM_URIAGE3, 50.0);
        Path f = tmp.resolve("2026年8月度 月次処理ファイル.xlsx");
        Files.write(f, new byte[] {1});
        CheckCResult match = CheckVerifyC.evaluate(f, monthly, 100, 80, 50, 0, 0.5);
        assertEquals(CheckCResult.MATCH, match.rows().get(0).judge());
        assertEquals(0, match.needCheckCount());

        CheckCResult expl = CheckVerifyC.evaluate(f, monthly, 100, 80, 70, 20, 0.5);
        assertEquals(CheckCResult.EXPLAINABLE, expl.rows().get(3).judge());
        assertEquals(0, expl.needCheckCount());

        CheckCResult need = CheckVerifyC.evaluate(f, monthly, 100, 80, 90, 5, 0.5);
        assertEquals(CheckCResult.NEED_CHECK, need.rows().get(3).judge());
        assertTrue(need.needCheckCount() >= 1);

        CheckCResult unread = CheckVerifyC.evaluate(f, Map.of(), 1, 1, 1, 0, 0.5);
        assertEquals(4, unread.needCheckCount());
        assertTrue(unread.rows().stream().allMatch(r -> CheckCResult.UNREADABLE.equals(r.judge())));
    }

    @Test
    void gokei1MatchesShisanEvenWhenShukeiDropsYs() throws Exception {
        Map<String, Double> monthly = new LinkedHashMap<>();
        monthly.put(CheckVerifyC.ITEM_JISSEKI1, 6_430_409.0);
        monthly.put(CheckVerifyC.ITEM_GOKEI1, 6_430_409.0);
        monthly.put(CheckVerifyC.ITEM_SHUKEI, 5_713_409.0);
        monthly.put(CheckVerifyC.ITEM_JISSEKI2, 6_430_409.0);
        monthly.put(CheckVerifyC.ITEM_URIAGE3, 6_430_409.0);
        Path f = tmp.resolve("2026年7月度 月次処理ファイル.xlsx");
        Files.write(f, new byte[] {1});
        CheckCResult r = CheckVerifyC.evaluate(f, monthly, 6_430_409.0, 6_430_409.0, 6_430_409.0, 0, 0.5);
        assertEquals(CheckCResult.MATCH, row(r, CheckVerifyC.ITEM_GOKEI1).judge());
        assertEquals(0, r.needCheckCount());
        CheckCResult.Row shukei = row(r, CheckVerifyC.ITEM_SHUKEI);
        assertEquals(CheckCResult.SHEET_INTERNAL, shukei.judge());
        assertTrue(shukei.note().contains("後加工集計"));
    }

    @Test
    void readMonthlyFileGokei1FromAtogakoSheet() throws Exception {
        Path f = tmp.resolve("2026年7月度 月次処理ファイル.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var gokei = wb.createSheet("Excel 後加工集計①");
            var r4 = gokei.createRow(3);
            r4.createCell(4).setCellValue("合計");
            var ys = gokei.createRow(4);
            ys.createCell(1).setCellValue("東レ　Ｙ.S");
            ys.createCell(4).setCellValue(717_000);
            var tvc = gokei.createRow(5);
            tvc.createCell(1).setCellValue("東レ T.V.C");
            tvc.createCell(4).setCellValue(4_059_163);
            var we = gokei.createRow(6);
            we.createCell(1).setCellValue("東レ W.E");
            we.createCell(4).setCellValue(1_654_246);
            var tot = gokei.createRow(7);
            tot.createCell(1).setCellValue("東レ 合計");
            tot.createCell(4).setCellValue(6_430_409);
            var sen = gokei.createRow(8);
            sen.createCell(4).setCellValue(6_430.409);
            var shukei = wb.createSheet("集計表");
            var sh = shukei.createRow(0);
            sh.createCell(0).setCellValue("区分");
            sh.createCell(1).setCellValue("加工金額");
            var sr = shukei.createRow(1);
            sr.createCell(0).setCellValue("東レ合計");
            sr.createCell(1).setCellValue(5_713_409);
            try (var out = Files.newOutputStream(f)) {
                wb.write(out);
            }
        }
        Map<String, Double> m = CheckVerifyC.readMonthlyFile(f);
        assertEquals(6_430_409.0, m.get(CheckVerifyC.ITEM_GOKEI1), 0.001);
        assertEquals(5_713_409.0, m.get(CheckVerifyC.ITEM_SHUKEI), 0.001);
    }

    private static CheckCResult.Row row(CheckCResult r, String item) {
        return r.rows().stream().filter(x -> item.equals(x.item())).findFirst().orElseThrow();
    }

    @Test
    void readMonthlyFileColumns() throws Exception {
        Path f = tmp.resolve("2026年8月度 月次処理ファイル.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var monthly = wb.createSheet("月次検証");
            var h = monthly.createRow(0);
            h.createCell(0).setCellValue("区分");
            h.createCell(1).setCellValue("月次実績1");
            h.createCell(2).setCellValue("月次実績表2");
            h.createCell(3).setCellValue("売上明細3");
            var t = monthly.createRow(1);
            t.createCell(0).setCellValue("東レ");
            t.createCell(1).setCellValue(10);
            t.createCell(2).setCellValue(20);
            t.createCell(3).setCellValue(30);
            var shukei = wb.createSheet("集計表");
            var sh = shukei.createRow(0);
            sh.createCell(0).setCellValue("区分");
            sh.createCell(1).setCellValue("加工金額");
            var sr = shukei.createRow(1);
            sr.createCell(0).setCellValue("東レ合計");
            sr.createCell(1).setCellValue(40);
            try (var out = Files.newOutputStream(f)) {
                wb.write(out);
            }
        }
        Map<String, Double> m = CheckVerifyC.readMonthlyFile(f);
        assertEquals(10.0, m.get(CheckVerifyC.ITEM_JISSEKI1), 0.001);
        assertEquals(20.0, m.get(CheckVerifyC.ITEM_JISSEKI2), 0.001);
        assertEquals(30.0, m.get(CheckVerifyC.ITEM_URIAGE3), 0.001);
        assertEquals(40.0, m.get(CheckVerifyC.ITEM_SHUKEI), 0.001);
    }
}
