package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class MasterDispatchSheetsSaveWriterTest {

    @TempDir Path tmp;

    @Test
    void save_backsUpThenWritesJsonAndWorkbook() throws Exception {
        Path json = tmp.resolve("master-dispatch-sheets.json");
        Path master = tmp.resolve("master.xlsx");
        Files.writeString(json, "{\"old\":true}\n");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("skills").createRow(0).createCell(0).setCellValue("旧");
            wb.createSheet("need");
            wb.createSheet("speed");
            wb.createSheet("組み合わせ表");
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
                                        "skills", List.of(List.of("工程名", "新")))));

        MasterDispatchSheetsSaveWriter.Result result =
                MasterDispatchSheetsSaveWriter.save(json, master, doc, Map.of());

        assertTrue(Files.isRegularFile(result.jsonBackup()));
        assertTrue(Files.isRegularFile(result.workbookBackup()));
        assertEquals("{\"old\":true}\n", Files.readString(result.jsonBackup()));
        MasterDispatchSheetsDocument loadedJson = MasterDispatchSheetsJsonStore.read(json);
        assertEquals("新", loadedJson.sheet("skills").rows().get(0).get(1));
        MasterDispatchSheetsDocument loadedWb =
                MasterDispatchSheetWorkbookImporter.importWorkbook(master, "KONAN");
        assertEquals("新", loadedWb.sheet("skills").rows().get(0).get(1));
    }
}
