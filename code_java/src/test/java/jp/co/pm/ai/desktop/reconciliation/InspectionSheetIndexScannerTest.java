package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class InspectionSheetIndexScannerTest {

    @Test
    void scansRecursivelyAndSkipsTempAndCsv(@TempDir Path tmp) throws Exception {
        Path month = tmp.resolve("2026年").resolve("9月");
        Files.createDirectories(month);
        Path xlsx = month.resolve("2026_C8-9(SEC済)完了.xlsx");
        writeKonan(xlsx, "C8-9", 46261);
        Files.writeString(month.resolve("ignore.csv"), "a,b");
        Files.writeString(month.resolve("~$lock.xlsx"), "x");

        InspectionSheetIndexScanner.Result first =
                InspectionSheetIndexScanner.scan(tmp, List.of(), null);
        assertEquals(1, first.rows().size());
        assertEquals("C8-9", first.rows().get(0).iraiNo());
        assertEquals(1, first.readExcelCount());

        InspectionSheetIndexScanner.Result second =
                InspectionSheetIndexScanner.scan(tmp, first.rows(), null);
        assertEquals(1, second.rows().size());
        assertEquals(0, second.readExcelCount());
    }

    @Test
    void scan_reportsWalkThenIndexProgress(@TempDir Path tmp) throws Exception {
        Path month = tmp.resolve("2026年").resolve("9月");
        Files.createDirectories(month);
        writeKonan(month.resolve("2026_C8-9(SEC済)完了.xlsx"), "C8-9", 46261);
        writeKonan(month.resolve("2026_C8-10(SEC済)完了.xlsx"), "C8-10", 46261);

        java.util.ArrayList<String> phases = new java.util.ArrayList<>();
        java.util.ArrayList<int[]> counts = new java.util.ArrayList<>();
        InspectionSheetIndexScanner.scan(
                tmp,
                List.of(),
                (phase, done, total) -> {
                    phases.add(phase);
                    counts.add(new int[] {done, total});
                });

        assertTrue(phases.contains(InspectionSheetIndexProgress.PHASE_WALK));
        assertTrue(phases.contains(InspectionSheetIndexProgress.PHASE_INDEX));
        int[] lastIndex = null;
        for (int i = 0; i < phases.size(); i++) {
            if (InspectionSheetIndexProgress.PHASE_INDEX.equals(phases.get(i))) {
                lastIndex = counts.get(i);
            }
        }
        org.junit.jupiter.api.Assertions.assertNotNull(lastIndex);
        assertEquals(2, lastIndex[0]);
        assertEquals(2, lastIndex[1]);
    }

    @Test
    void scan_mergesTwoRoots(@TempDir Path tmp) throws Exception {
        Path unc = tmp.resolve("unc");
        Path box = tmp.resolve("box");
        Files.createDirectories(unc);
        Files.createDirectories(box);
        writeKonan(unc.resolve("2026_C8-9(SEC済)完了.xlsx"), "C8-9", 46261);
        writeKonan(box.resolve("2026_GB60804(ｽﾗｲｽ済)完了.xlsx"), "GB60804", 46261);

        InspectionSheetIndexScanner.Result merged =
                InspectionSheetIndexScanner.scan(List.of(unc, box), List.of(), null);
        assertEquals(2, merged.rows().size());
        assertTrue(merged.rows().stream().anyMatch(r -> "C8-9".equals(r.iraiNo())));
        assertTrue(merged.rows().stream().anyMatch(r -> "GB60804".equals(r.iraiNo())));
    }

    @Test
    void scan_manyFiles_keepsPathOrder(@TempDir Path tmp) throws Exception {
        Path month = tmp.resolve("2026年").resolve("9月");
        Files.createDirectories(month);
        String[] irais = {"A1-1", "A1-2", "B2-1", "C3-1", "D4-1", "E5-1", "F6-1", "G7-1"};
        for (String irai : irais) {
            writeKonan(month.resolve("2026_" + irai + "(SEC済)完了.xlsx"), irai, 46261);
        }
        InspectionSheetIndexScanner.Result result =
                InspectionSheetIndexScanner.scan(tmp, List.of(), null);
        assertEquals(irais.length, result.rows().size());
        assertEquals(irais.length, result.readExcelCount());
        List<String> got = result.rows().stream().map(r -> r.iraiNo()).toList();
        assertEquals(List.of(irais), got);
    }

    private static void writeKonan(Path file, String irai, double serial) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sh = wb.createSheet("検査表");
            var r2 = sh.createRow(1);
            r2.createCell(26).setCellValue("加工日");
            r2.createCell(28).setCellValue(serial);
            var r3 = sh.createRow(2);
            r3.createCell(0).setCellValue("加工依頼№");
            r3.createCell(3).setCellValue(irai);
            try (var out = Files.newOutputStream(file)) {
                wb.write(out);
            }
        }
    }
}
