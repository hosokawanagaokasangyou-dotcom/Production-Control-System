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
