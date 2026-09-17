package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class VerifyEngineSkipBTest {

    @TempDir
    Path tmp;

    @Test
    @DisplayName("③の対象年月が違うときは検証BをスキップしてAとExcelを出す")
    void skipBWhenAladdinYmMismatches() throws Exception {
        KouchinPaths paths = setup(tmp, "2026年06月", true);
        VerifyResult result = VerifyEngine.run(FactoryProfile.of(FactoryId.KOKUBU), paths, 0.5);
        assertTrue(result.skippedB(), "検証Bをスキップすること");
        assertTrue(result.recordsB().isEmpty(), "B明細は空");
        assertTrue(result.recordsA().stream().anyMatch(r -> "191352R".equals(r.keiyaku())), "検証Aは実行");
        assertTrue(result.warnings().stream().anyMatch(w -> w.contains("【検証Bスキップ】")));
        assertWorkbookShowsSkip(result);
    }

    @Test
    @DisplayName("③ファイル自体が無いときも検証BをスキップしてExcelを出す")
    void skipBWhenAladdinFileMissing() throws Exception {
        KouchinPaths paths = setup(tmp, "2026年07月", true);
        Files.deleteIfExists(paths.kokubuAladdinDir().resolve("依頼NO別問合せ_20260731_120000.xlsx"));
        VerifyResult result = VerifyEngine.run(FactoryProfile.of(FactoryId.KOKUBU), paths, 0.5);
        assertTrue(result.skippedB());
        assertTrue(result.recordsB().isEmpty());
        assertTrue(result.recordsA().stream().anyMatch(r -> "191352R".equals(r.keiyaku())));
        assertWorkbookShowsSkip(result);
    }

    @Test
    @DisplayName("③に加工金額が無いときは検証BをスキップしてExcelを出す")
    void skipBWhenAladdinHasNoKakouKingaku() throws Exception {
        KouchinPaths paths = setup(tmp, "2026年07月", false);
        VerifyResult result = VerifyEngine.run(FactoryProfile.of(FactoryId.KOKUBU), paths, 0.5);
        assertTrue(result.skippedB());
        assertTrue(result.recordsB().isEmpty());
        assertTrue(result.recordsA().stream().anyMatch(r -> "191352R".equals(r.keiyaku())));
        assertTrue(result.warnings().stream().anyMatch(w -> w.contains("【検証Bスキップ】")));
        assertWorkbookShowsSkip(result);
    }

    @Test
    @DisplayName("③に対象月データがあるときは検証Bを実行する")
    void runBWhenAladdinCoversTargetMonth() throws Exception {
        KouchinPaths paths = setup(tmp, "2026年07月", true);
        VerifyResult result = VerifyEngine.run(FactoryProfile.of(FactoryId.KOKUBU), paths, 0.5);
        assertFalse(result.skippedB());
        assertTrue(result.recordsB().stream().anyMatch(r -> "C-8".equals(r.irai())));
        assertTrue(result.warnings().stream().noneMatch(w -> w.contains("【検証Bスキップ】")));
    }

    private static void assertWorkbookShowsSkip(VerifyResult result) throws Exception {
        try (XSSFWorkbook wb = ResultExcelExporter.buildWorkbook(result, null, null)) {
            XSSFSheet summary = wb.getSheet("サマリ");
            assertTrue(sheetContains(summary, "【検証Bスキップ】"), "サマリ先頭付近に検証Bスキップ警告");
            assertTrue(sheetContains(summary, "スキップ"), "KPIまたは検証B欄にスキップ");
            XSSFSheet b = wb.getSheet("検証B_依頼NO(②vs③)");
            assertTrue(sheetContains(b, "スキップ"), "検証Bシートにスキップ表示");
        }
    }

    private static boolean sheetContains(XSSFSheet ws, String needle) {
        if (ws == null) {
            return false;
        }
        for (Row row : ws) {
            for (Cell cell : row) {
                String v = cell.toString();
                if (v != null && v.contains(needle)) {
                    return true;
                }
            }
        }
        return false;
    }

    private static KouchinPaths setup(Path tmp, String taishoYm, boolean withKakou) throws Exception {
        Path csvDir = tmp.resolve("csv");
        Path nagaokaDir = tmp.resolve("nagaoka");
        Path aladdinDir = tmp.resolve("aladdin");
        Files.createDirectories(csvDir);
        Files.createDirectories(nagaokaDir);
        Files.createDirectories(aladdinDir);

        String csv = String.join("\n",
                "入庫場所,生産場所,発注No.,入庫日,品番,品名カナ,数量,品名,単価,金額,備考1,備考2",
                "A010,N122B,191-352R,250705,X1,ｶﾅ,10,,1000,10000,,",
                "");
        Files.write(csvDir.resolve("RVSHEET202607.csv"), csv.getBytes(Charset.forName("windows-31j")));

        Path nagaoka = nagaokaDir.resolve("後加工工賃明細（2026年7月度).xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet ws = wb.createSheet("東レまとめ");
            cell(ws, 2, 2, "契約No.");
            cell(ws, 3, 26, "合計");
            cell(ws, 5, 0, "C");
            cell(ws, 5, 1, 8);
            cell(ws, 5, 2, "191352R");
            cell(ws, 5, 26, 10000);
            try (var os = Files.newOutputStream(nagaoka)) {
                wb.write(os);
            }
        }

        Path aladdin = aladdinDir.resolve("依頼NO別問合せ_20260731_120000.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet ws = wb.createSheet("問合せ");
            cell(ws, 0, 0, "対象年月 : " + taishoYm);
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
            try (var os = Files.newOutputStream(aladdin)) {
                wb.write(os);
            }
        }

        return new KouchinPaths(
                csvDir, nagaokaDir, aladdinDir,
                tmp.resolve("shisan"), tmp.resolve("konan3"), tmp.resolve("monthly"),
                tmp, tmp);
    }

    private static void cell(XSSFSheet ws, int r, int c, String v) {
        Row row = ws.getRow(r) == null ? ws.createRow(r) : ws.getRow(r);
        row.createCell(c).setCellValue(v);
    }

    private static void cell(XSSFSheet ws, int r, int c, double v) {
        Row row = ws.getRow(r) == null ? ws.createRow(r) : ws.getRow(r);
        row.createCell(c).setCellValue(v);
    }
}
