package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class Source3TargetMonthCheckTest {

    @TempDir
    Path tmp;

    private static final YearMonthKey JUL = new YearMonthKey(2026, 7);

    @Test
    @DisplayName("対象月データが無いときの備考は警告文")
    void warningNoteNamesTargetMonth() {
        String note = Source3TargetMonthCheck.warningNote(JUL, "加工金額行が0件");
        assertTrue(note.contains("対象月（2026年7月度）のデータがありません"), note);
        assertTrue(note.contains("加工金額行が0件"), note);
        assertTrue(Source3TargetMonthCheck.isWarning(note));
        assertFalse(Source3TargetMonthCheck.isWarning(""));
        assertFalse(Source3TargetMonthCheck.isWarning("見つかりません"));
        assertEquals("pm-kouchin-note-warn", Source3TargetMonthCheck.NOTE_CSS);
    }

    @Test
    @DisplayName("③の対象年月が違う・加工金額が無いときは警告備考")
    void uncoveredFileGetsWarningNote() throws Exception {
        Path mismatch = writeAladdin("依頼NO別問合せ_wrong.xlsx", "2026年06月", true);
        Path empty = writeAladdin("依頼NO別問合せ_empty.xlsx", "2026年07月", false);
        Path ok = writeAladdin("依頼NO別問合せ_ok.xlsx", "2026年07月", true);
        FactoryProfile kokubu = FactoryProfile.of(FactoryId.KOKUBU);
        String wrong = Source3TargetMonthCheck.noteIfUncovered(mismatch, JUL, kokubu);
        String none = Source3TargetMonthCheck.noteIfUncovered(empty, JUL, kokubu);
        String covered = Source3TargetMonthCheck.noteIfUncovered(ok, JUL, kokubu);
        assertTrue(Source3TargetMonthCheck.isWarning(wrong), wrong);
        assertTrue(wrong.contains("シート対象年月"), wrong);
        assertTrue(Source3TargetMonthCheck.isWarning(none), none);
        assertTrue(none.contains("加工金額行が0件"), none);
        assertEquals("", covered);
    }

    @Test
    @DisplayName("検出表の③行は対象月なしなら備考に警告を載せる")
    void scanPutsWarningOnUncoveredSource3() throws Exception {
        Path csvDir = tmp.resolve("csv");
        Path aladdinDir = tmp.resolve("aladdin");
        Files.createDirectories(csvDir);
        Files.createDirectories(aladdinDir);
        Files.writeString(csvDir.resolve("RVSHEET202607.csv"), "入庫場所\n");
        writeAladdin(aladdinDir.resolve("依頼NO別問合せ_20260731_120000.xlsx"), "2026年07月", false);
        KouchinPaths paths = new KouchinPaths(
                csvDir,
                tmp.resolve("nagaoka"),
                aladdinDir,
                tmp.resolve("shisan"),
                tmp.resolve("konan3"),
                tmp.resolve("monthly"),
                tmp.resolve("out"),
                tmp.resolve("judge"));
        KouchinDiscovery.Row row3 = KouchinDiscovery.scan(FactoryId.KOKUBU, paths).stream()
                .filter(r -> KouchinDiscovery.ROLE_3.equals(r.role()))
                .findFirst()
                .orElseThrow();
        assertFalse(row3.missing());
        assertTrue(Source3TargetMonthCheck.isWarning(row3.note()), row3.note());
    }

    private Path writeAladdin(String name, String taishoYm, boolean withKakou) throws Exception {
        return writeAladdin(tmp.resolve(name), taishoYm, withKakou);
    }

    private static Path writeAladdin(Path path, String taishoYm, boolean withKakou) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet ws = wb.createSheet("問合せ");
            ws.createRow(0).createCell(0).setCellValue("対象年月 : " + taishoYm);
            Row h = ws.createRow(1);
            h.createCell(0).setCellValue("依頼NO");
            h.createCell(1).setCellValue("項目");
            h.createCell(2).setCellValue("--合計--");
            if (withKakou) {
                Row d = ws.createRow(2);
                d.createCell(0).setCellValue("C-8");
                d.createCell(1).setCellValue("加工金額");
                d.createCell(2).setCellValue(10000);
            }
            try (var os = Files.newOutputStream(path)) {
                wb.write(os);
            }
        }
        return path;
    }
}
