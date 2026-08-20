package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class MasterDispatchSheetWorkbookExporterTest {

    @TempDir Path tmp;

    @Test
    void writeBack_updatesFourSheetsPreservesOthersAndClearsLeftoverCells() throws Exception {
        Path master = tmp.resolve("master.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var skills = wb.createSheet("skills");
            skills.createRow(0).createCell(0).setCellValue("工程名");
            skills.getRow(0).createCell(1).setCellValue("巻返し");
            skills.createRow(1).createCell(0).setCellValue("旧メンバー");
            skills.getRow(1).createCell(1).setCellValue("OP9");
            skills.createRow(2).createCell(0).setCellValue("削除される行");

            var need = wb.createSheet("need");
            need.createRow(0).createCell(0).setCellValue("工程名");

            var speed = wb.createSheet("speed");
            speed.createRow(0).createCell(0).setCellValue("工程名");
            speed.createRow(3).createCell(3).setCellValue(1.0);

            var combo = wb.createSheet("組み合わせ表");
            combo.createRow(0).createCell(0).setCellValue("組み合わせ行ID");

            wb.createSheet("組み合わせ表20260404").createRow(0).createCell(0).setCellValue("keep");

            try (OutputStream os = Files.newOutputStream(master)) {
                wb.write(os);
            }
        }

        MasterDispatchSheetsDocument doc =
                new MasterDispatchSheetsDocument(
                        1,
                        "KONAN",
                        master.toString(),
                        "2026-08-21T07:00:00+09:00",
                        Map.of(
                                "skills",
                                new MasterDispatchSheetsDocument.SheetGrid(
                                        "skills",
                                        List.of(
                                                List.of("工程名", "巻返し"),
                                                List.of("山田", "OP1"))),
                                "need",
                                new MasterDispatchSheetsDocument.SheetGrid(
                                        "need", List.of(List.of("工程名", "2"))),
                                "speed",
                                new MasterDispatchSheetsDocument.SheetGrid(
                                        "speed",
                                        List.of(
                                                List.of("工程名", ""),
                                                List.of("", ""),
                                                List.of("", ""),
                                                List.of("", "", "", "12.5"))),
                                "teamCombinations",
                                new MasterDispatchSheetsDocument.SheetGrid(
                                        "組み合わせ表",
                                        List.of(
                                                List.of("組み合わせ行ID", "工程名"),
                                                List.of("1", "巻返し")))));

        MasterDispatchSheetWorkbookExporter.writeBack(master, doc, Map.of());

        MasterDispatchSheetsDocument loaded =
                MasterDispatchSheetWorkbookImporter.importWorkbook(master, "KONAN");
        assertEquals(List.of("工程名", "巻返し"), loaded.sheet("skills").rows().get(0));
        assertEquals(List.of("山田", "OP1"), loaded.sheet("skills").rows().get(1));
        assertEquals(2, loaded.sheet("skills").rows().size());
        assertEquals("2", loaded.sheet("need").rows().get(0).get(1));
        assertEquals("12.5", loaded.sheet("speed").rows().get(3).get(3));
        assertEquals("巻返し", loaded.sheet("teamCombinations").rows().get(1).get(1));

        try (Workbook wb = WorkbookFactory.create(master.toFile())) {
            Sheet extra = wb.getSheet("組み合わせ表20260404");
            assertEquals("keep", extra.getRow(0).getCell(0).getStringCellValue());
            Sheet skills = wb.getSheet("skills");
            Row leftover = skills.getRow(2);
            if (leftover != null) {
                Cell c0 = leftover.getCell(0);
                assertTrue(c0 == null || new DataFormatter().formatCellValue(c0).isBlank());
            }
        }
    }

    @Test
    void writeBack_missingWorkbook_throws() {
        Path missing = tmp.resolve("no-master.xlsm");
        IOException ex =
                assertThrows(
                        IOException.class,
                        () ->
                                MasterDispatchSheetWorkbookExporter.writeBack(
                                        missing, MasterDispatchSheetsDocument.empty("KONAN"), Map.of()));
        assertTrue(ex.getMessage().contains("master"));
    }
}
