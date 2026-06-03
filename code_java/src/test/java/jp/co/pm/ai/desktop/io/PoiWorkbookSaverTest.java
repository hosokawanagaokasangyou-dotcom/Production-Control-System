package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertFalse;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

class PoiWorkbookSaverTest {

    /**
     * 転記と同様に数式セルを removeCell + createCell で値に置き換えたあと保存しても
     * calcChain.xml を残さない（Excel 修復ダイアログの原因になりやすい）。
     */
    @Test
    void write_omitsCalcChainAfterFormulaCellReplacedLikeJuchuTransfer() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("受注ﾌｧｲﾙ");
            XSSFRow row = sheet.createRow(3);
            XSSFCell formulaCell = row.createCell(2);
            formulaCell.setCellFormula("A1+1");

            row.removeCell(formulaCell);
            row.createCell(2).setCellValue("replaced");

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            PoiWorkbookSaver.write(wb, out);

            assertFalse(zipContainsEntry(out.toByteArray(), "xl/calcChain.xml"));
        }
    }

    private static boolean zipContainsEntry(byte[] zipBytes, String entryName) throws Exception {
        try (ZipInputStream zis = new ZipInputStream(new ByteArrayInputStream(zipBytes))) {
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                if (entryName.equals(entry.getName())) {
                    return true;
                }
            }
        }
        return false;
    }
}
