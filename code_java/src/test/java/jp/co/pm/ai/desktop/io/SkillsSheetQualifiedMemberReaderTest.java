package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class SkillsSheetQualifiedMemberReaderTest {

    @Test
    void readQualifiedMemberDisplayNames_returnsOnlyOpAndAsForProcessMachine(
            @TempDir Path root) throws Exception {
        Path master = root.resolve("master.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sheet = wb.createSheet("skills");
            sheet.createRow(0).createCell(0).setCellValue("工程名");
            sheet.getRow(0).createCell(1).setCellValue("ラミネート");
            sheet.getRow(0).createCell(2).setCellValue("ラミネート");
            sheet.createRow(1).createCell(0).setCellValue("機械名");
            sheet.getRow(1).createCell(1).setCellValue("1号機");
            sheet.getRow(1).createCell(2).setCellValue("2号機");

            var yamada = sheet.createRow(2);
            yamada.createCell(0).setCellValue("山田");
            yamada.createCell(1).setCellValue("OP1");
            yamada.createCell(2).setCellValue("AS1");

            var sato = sheet.createRow(3);
            sato.createCell(0).setCellValue("佐藤");
            sato.createCell(1).setCellValue("AS 2");

            var suzuki = sheet.createRow(4);
            suzuki.createCell(0).setCellValue("鈴木");
            suzuki.createCell(1).setCellValue("見習");

            try (var out = Files.newOutputStream(master)) {
                wb.write(out);
            }
        }

        assertEquals(
                List.of("山田", "佐藤"),
                SkillsSheetMemberReader.readQualifiedMemberDisplayNames(
                        master, "ラミネート", "1号機"));
        assertEquals(
                List.of("山田"),
                SkillsSheetMemberReader.readQualifiedMemberDisplayNames(
                        master, "ラミネート", "2号機"));
    }

    @Test
    void readQualifiedMemberDisplayNames_recognizesLegacySingleHeaderDeterministically(
            @TempDir Path root) throws Exception {
        Path master = root.resolve("legacy-master.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sheet = wb.createSheet("skills");
            var header = sheet.createRow(0);
            header.createCell(0).setCellValue("メンバー");
            header.createCell(1).setCellValue("ラミネート+1号機");
            header.createCell(2).setCellValue("ラミネート+2号機");

            var yamada = sheet.createRow(1);
            yamada.createCell(0).setCellValue("山田");
            yamada.createCell(1).setCellValue("OP1");

            var sato = sheet.createRow(2);
            sato.createCell(0).setCellValue("佐藤");
            sato.createCell(1).setCellValue("AS2");

            var suzuki = sheet.createRow(3);
            suzuki.createCell(0).setCellValue("鈴木");
            suzuki.createCell(1).setCellValue("研修中");

            try (var out = Files.newOutputStream(master)) {
                wb.write(out);
            }
        }

        assertEquals(
                List.of("山田", "佐藤"),
                SkillsSheetMemberReader.readQualifiedMemberDisplayNames(
                        master, "ラミネート", "1号機"));
        assertEquals(
                List.of("山田", "佐藤", "鈴木"),
                SkillsSheetMemberReader.readMemberDisplayNames(master));
    }
}
