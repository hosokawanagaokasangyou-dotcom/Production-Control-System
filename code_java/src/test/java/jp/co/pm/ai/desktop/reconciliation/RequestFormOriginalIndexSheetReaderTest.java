package jp.co.pm.ai.desktop.reconciliation;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class RequestFormOriginalIndexSheetReaderTest {

    @Test
    void read_parsesRowsByIraiNo() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet index = wb.createSheet("目次");
            index.createRow(1).createCell(0).setCellValue("加工依頼NO");
            var row1 = index.createRow(2);
            row1.createCell(0).setCellValue("T6-20");
            row1.createCell(8).setCellValue("6/9");
            row1.createCell(9).setCellValue("6/10");
            row1.createCell(10).setCellValue("6/22");
            row1.createCell(11).setCellValue("納期メモ");
            row1.createCell(13).setCellValue("185821Z");
            row1.createCell(14).setCellValue("契約メモ");
            var row2 = index.createRow(3);
            row2.createCell(0).setCellValue("T6-21");
            row2.createCell(9).setCellValue("6/11");

            Map<String, RequestFormOriginalIndexSheetReader.IndexEntry> map =
                    RequestFormOriginalIndexSheetReader.read(index);

            assertEquals(2, map.size());
            RequestFormOriginalIndexSheetReader.IndexEntry t620 =
                    map.get(JuchuTransferValueNormalizer.normalizeKey("T6-20"));
            assertTrue(t620 != null);
            assertEquals("T6-20", t620.iraiNo());
            assertEquals("6/9", t620.responseDate());
            assertEquals("6/10", t620.inputDate());
            assertEquals("6/22", t620.deliveryDate());
            assertEquals("納期メモ", t620.deliveryRemarks());
            assertEquals("185821Z", t620.contractNo());
            assertEquals("契約メモ", t620.contractRemarks());
        }
    }
}
