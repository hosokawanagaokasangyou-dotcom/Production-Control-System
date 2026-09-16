package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/**
 * Discovery → Engine → Writer の @TempDir スモーク。実 UNC は使わない。
 */
class VerifyPipelineSmokeTest {

    @TempDir Path tmp;

    @Test
    void discoveryEngineExporterDualWrite() throws Exception {
        Path csvDir = tmp.resolve("csv");
        Path nagaokaDir = tmp.resolve("nagaoka");
        Path aladdinDir = tmp.resolve("aladdin");
        Path out1 = tmp.resolve("out1");
        Path out2 = tmp.resolve("out2");
        Files.createDirectories(csvDir);
        Files.createDirectories(nagaokaDir);
        Files.createDirectories(aladdinDir);
        Files.createDirectories(out1);
        Files.createDirectories(out2);

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
            cell(ws, 0, 0, "対象年月 : 2026年07月");
            Row h = ws.createRow(1);
            h.createCell(0).setCellValue("依頼NO");
            h.createCell(1).setCellValue("項目");
            h.createCell(2).setCellValue("--合計--");
            Row d = ws.createRow(2);
            d.createCell(0).setCellValue("C-8");
            d.createCell(1).setCellValue("加工金額");
            d.createCell(2).setCellValue(10000);
            try (var os = Files.newOutputStream(aladdin)) {
                wb.write(os);
            }
        }

        Path foundCsv = FileDiscovery.findTorayCsv(csvDir);
        assertEquals("RVSHEET202607.csv", foundCsv.getFileName().toString());

        KouchinPaths paths = new KouchinPaths(
                csvDir, nagaokaDir, aladdinDir,
                tmp.resolve("shisan"), tmp.resolve("konan3"), tmp.resolve("monthly"),
                tmp, tmp);
        VerifyResult result = VerifyEngine.run(FactoryProfile.of(FactoryId.KOKUBU), paths, 0.5);
        assertEquals(new YearMonthKey(2026, 7), result.targetYm());
        assertTrue(result.recordsA().stream().anyMatch(r -> "191352R".equals(r.keiyaku())));

        Path x1 = out1.resolve("検証結果_国分工場_smoke.xlsx");
        Path x2 = out2.resolve("検証結果_国分工場_smoke.xlsx");
        try (var wb = ResultExcelExporter.buildWorkbook(result, null)) {
            DualWriteFiles.WriteOutcome o = DualWriteFiles.writeWorkbook(wb, List.of(x1, x2), Map.of());
            assertEquals(2, o.succeeded().size());
            assertTrue(Files.size(x1) > 0);
            assertTrue(Files.size(x2) > 0);
        }
        ResultArchive.archiveOldVerifyResults(out1, List.of(x1));
        assertTrue(Files.isRegularFile(x1));
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
