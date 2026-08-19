package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class MasterDispatchSheetWorkbookImporterTest {

    @TempDir Path tmp;

    @Test
    void importsFourSheetsAsStringGridAndTrimsTrailingEmpty() throws Exception {
        Path master = tmp.resolve("master.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var skills = wb.createSheet("skills");
            skills.createRow(0).createCell(0).setCellValue("工程名");
            skills.getRow(0).createCell(1).setCellValue("巻返し");
            skills.createRow(1).createCell(0).setCellValue("機械名");
            skills.getRow(1).createCell(1).setCellValue("機1");
            skills.createRow(2).createCell(0).setCellValue("山田");
            skills.getRow(2).createCell(1).setCellValue("OP1");
            skills.createRow(3).createCell(0).setCellValue("");
            skills.getRow(3).createCell(1).setCellValue("");

            var need = wb.createSheet("need");
            need.createRow(0).createCell(0).setCellValue("工程名");
            need.getRow(0).createCell(1).setCellValue("巻返し");

            var speed = wb.createSheet("speed");
            speed.createRow(0).createCell(0).setCellValue("工程名");
            speed.createRow(3).createCell(3).setCellValue(12.5);
            speed.createRow(4).createCell(3).setCellValue(0.8);

            var combo = wb.createSheet("組み合わせ表");
            combo.createRow(0).createCell(0).setCellValue("組み合わせ行ID");
            combo.getRow(0).createCell(1).setCellValue("工程名");
            combo.createRow(1).createCell(0).setCellValue("1");
            combo.getRow(1).createCell(1).setCellValue("巻返し");

            wb.createSheet("組み合わせ表20260404");

            try (OutputStream os = Files.newOutputStream(master)) {
                wb.write(os);
            }
        }

        MasterDispatchSheetsDocument doc =
                MasterDispatchSheetWorkbookImporter.importWorkbook(master, "KONAN");

        assertEquals(1, doc.schemaVersion());
        assertEquals("KONAN", doc.factorySite());
        assertEquals(master.toAbsolutePath().normalize().toString(), doc.sourceWorkbook());
        assertEquals("skills", doc.sheet("skills").sheetName());
        assertEquals(List.of("工程名", "巻返し"), doc.sheet("skills").rows().get(0));
        assertEquals(List.of("山田", "OP1"), doc.sheet("skills").rows().get(2));
        assertEquals(3, doc.sheet("skills").rows().size());
        assertEquals("need", doc.sheet("need").sheetName());
        assertEquals("speed", doc.sheet("speed").sheetName());
        assertEquals("12.5", doc.sheet("speed").rows().get(3).get(3));
        assertEquals("組み合わせ表", doc.sheet("teamCombinations").sheetName());
        assertEquals("巻返し", doc.sheet("teamCombinations").rows().get(1).get(1));
    }

    @Test
    void missingSheetsBecomeEmptyRows() throws Exception {
        Path master = tmp.resolve("master.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("skills");
            try (OutputStream os = Files.newOutputStream(master)) {
                wb.write(os);
            }
        }
        MasterDispatchSheetsDocument doc =
                MasterDispatchSheetWorkbookImporter.importWorkbook(master, "KOKUBU");
        assertEquals("KOKUBU", doc.factorySite());
        assertTrue(doc.sheet("need").rows().isEmpty());
        assertTrue(doc.sheet("speed").rows().isEmpty());
        assertTrue(doc.sheet("teamCombinations").rows().isEmpty());
        assertTrue(doc.sheet("skills").rows().isEmpty());
    }

    @Test
    void jsonStoreRoundTripsDocument(@TempDir Path dir) throws Exception {
        MasterDispatchSheetsDocument original =
                new MasterDispatchSheetsDocument(
                        1,
                        "KONAN",
                        "C:\\master.xlsm",
                        "2026-08-20T07:00:00+09:00",
                        java.util.Map.of(
                                "skills",
                                new MasterDispatchSheetsDocument.SheetGrid(
                                        "skills", List.of(List.of("工程名", "巻返し"))),
                                "need",
                                new MasterDispatchSheetsDocument.SheetGrid("need", List.of()),
                                "speed",
                                new MasterDispatchSheetsDocument.SheetGrid("speed", List.of()),
                                "teamCombinations",
                                new MasterDispatchSheetsDocument.SheetGrid(
                                        "組み合わせ表", List.of())));
        Path json = dir.resolve("master-dispatch-sheets.json");
        MasterDispatchSheetsJsonStore.write(json, original);
        MasterDispatchSheetsDocument loaded = MasterDispatchSheetsJsonStore.read(json);
        assertEquals(original.schemaVersion(), loaded.schemaVersion());
        assertEquals(original.factorySite(), loaded.factorySite());
        assertEquals(original.sourceWorkbook(), loaded.sourceWorkbook());
        assertEquals(original.sheet("skills").rows(), loaded.sheet("skills").rows());
        assertEquals("組み合わせ表", loaded.sheet("teamCombinations").sheetName());
    }
}
