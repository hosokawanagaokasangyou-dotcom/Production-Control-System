package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class SkillsSheetEquipmentListReaderTest {

    @Test
    void readsTwoHeaderEquipmentCombos(@TempDir Path tmp) throws Exception {
        Path master = tmp.resolve("master.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sh = wb.createSheet("skills");
            sh.createRow(0).createCell(1).setCellValue("工程P");
            sh.createRow(1).createCell(1).setCellValue("機械M");
            sh.createRow(2).createCell(0).setCellValue("オペレータ1");
            try (OutputStream os = Files.newOutputStream(master)) {
                wb.write(os);
            }
        }
        assertEquals(List.of("工程P+機械M"), SkillsSheetEquipmentListReader.readEquipmentProcPlusMachineCombos(master));
    }
}
