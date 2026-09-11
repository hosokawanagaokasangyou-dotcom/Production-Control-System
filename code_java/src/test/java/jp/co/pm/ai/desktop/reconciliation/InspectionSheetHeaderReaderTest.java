package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class InspectionSheetHeaderReaderTest {

    @Test
    void readsKonanInspectionSheetLayout(@TempDir Path tmp) throws Exception {
        Path file = tmp.resolve("2026_C8-9(SEC済)完了.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sh = wb.createSheet("検査表");
            var r2 = sh.createRow(1);
            r2.createCell(26).setCellValue("加工日カコウビ");
            r2.createCell(28).setCellValue(46261);
            var r3 = sh.createRow(2);
            r3.createCell(0).setCellValue("加工依頼№カコウイライ");
            r3.createCell(3).setCellValue("C8-9");
            try (var out = Files.newOutputStream(file)) {
                wb.write(out);
            }
        }
        InspectionSheetHeaderReader.Header header = InspectionSheetHeaderReader.read(file);
        assertEquals("C8-9", header.iraiNo());
        assertEquals(LocalDate.of(2026, 8, 27), header.processingDate());
    }

    @Test
    void readsOldKonanFirstSheetWhenNotNamedKensa(@TempDir Path tmp) throws Exception {
        Path file = tmp.resolve("C10-10（SEC済）完了.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sh = wb.createSheet("①SEC");
            var r2 = sh.createRow(1);
            r2.createCell(26).setCellValue("加工日");
            r2.createCell(28).setCellValue(44117);
            var r3 = sh.createRow(2);
            r3.createCell(0).setCellValue("加工依頼№");
            r3.createCell(3).setCellValue("C10-10");
            wb.createSheet("②ｽﾘｯﾄ");
            try (var out = Files.newOutputStream(file)) {
                wb.write(out);
            }
        }
        InspectionSheetHeaderReader.Header header = InspectionSheetHeaderReader.read(file);
        assertEquals("C10-10", header.iraiNo());
        assertEquals(InspectionSheetProcessingDate.fromExcelSerial(44117).orElseThrow(), header.processingDate());
    }

    @Test
    void readsKokubuEcLayout_valueInC2_rangeDate(@TempDir Path tmp) throws Exception {
        Path file = tmp.resolve("C1-4#product(EC.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sh = wb.createSheet("C1-4");
            var r2 = sh.createRow(1);
            r2.createCell(1).setCellValue("依頼No．イライ");
            r2.createCell(2).setCellValue("C1-4");
            r2.createCell(5).setCellValue("加工日： 2026/1/9 ～ 1/13");
            try (var out = Files.newOutputStream(file)) {
                wb.write(out);
            }
        }
        InspectionSheetHeaderReader.Header header = InspectionSheetHeaderReader.read(file);
        assertEquals("C1-4", header.iraiNo());
        assertEquals(LocalDate.of(2026, 1, 9), header.processingDate());
    }

    @Test
    void readsKokubuLacLayout_valueInD2(@TempDir Path tmp) throws Exception {
        Path file = tmp.resolve("C7-35#product(LAC.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sh = wb.createSheet("C7-35");
            var r2 = sh.createRow(1);
            r2.createCell(1).setCellValue("依頼No.イライ");
            r2.createCell(3).setCellValue("C7-35");
            var r3 = sh.createRow(2);
            r3.createCell(1).setCellValue("加工日： 2026/7/16 ～ 7/23");
            try (var out = Files.newOutputStream(file)) {
                wb.write(out);
            }
        }
        InspectionSheetHeaderReader.Header header = InspectionSheetHeaderReader.read(file);
        assertEquals("C7-35", header.iraiNo());
        assertEquals(LocalDate.of(2026, 7, 16), header.processingDate());
    }

    @Test
    void readsKokubuSliceFullwidthNo(@TempDir Path tmp) throws Exception {
        Path file = tmp.resolve("C1-1#product(スライス.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sh = wb.createSheet("C1-1");
            var r2 = sh.createRow(1);
            r2.createCell(1).setCellValue("依頼Ｎｏ：");
            r2.createCell(3).setCellValue("C1-1");
            r2.createCell(7).setCellValue("加工日： 2026/1/7");
            try (var out = Files.newOutputStream(file)) {
                wb.write(out);
            }
        }
        InspectionSheetHeaderReader.Header header = InspectionSheetHeaderReader.read(file);
        assertEquals("C1-1", header.iraiNo());
        assertEquals(LocalDate.of(2026, 1, 7), header.processingDate());
    }

    @Test
    void fallsBackToFileNameWhenHeaderMissing(@TempDir Path tmp) throws Exception {
        Path file = tmp.resolve("2026_C8-9(SEC済)完了.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("空");
            try (var out = Files.newOutputStream(file)) {
                wb.write(out);
            }
        }
        InspectionSheetHeaderReader.Header header = InspectionSheetHeaderReader.read(file);
        assertEquals("C8-9", header.iraiNo());
        assertNull(header.processingDate());
    }

    @Test
    void readsTpiIraiWithSpace(@TempDir Path tmp) throws Exception {
        Path file = tmp.resolve("TPI 1-1#product(融着.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sh = wb.createSheet("TPI 1-1");
            var r2 = sh.createRow(1);
            r2.createCell(1).setCellValue("依頼No．");
            r2.createCell(3).setCellValue("TPI 1-1");
            r2.createCell(5).setCellValue("加工日： 2026/1/8");
            try (var out = Files.newOutputStream(file)) {
                wb.write(out);
            }
        }
        InspectionSheetHeaderReader.Header header = InspectionSheetHeaderReader.read(file);
        assertEquals("TPI 1-1", header.iraiNo());
        assertEquals(LocalDate.of(2026, 1, 8), header.processingDate());
    }
}
