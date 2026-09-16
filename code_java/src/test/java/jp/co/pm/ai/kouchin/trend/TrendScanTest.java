package jp.co.pm.ai.kouchin.trend;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.FileTime;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.kouchin.verify.YearMonthKey;

class TrendScanTest {

    @TempDir Path tmp;

    @Test
    @DisplayName("後加工工賃明細* の yyyy年m月度から月を拾い、同月は更新日時が新しい方")
    void listKokubuPicksLatestPerYm() throws Exception {
        Path d = tmp.resolve("長岡");
        Files.createDirectories(d);
        Files.write(d.resolve("後加工工賃明細（2026年7月度)V01.xlsx"), new byte[] {'P', 'K'});
        Path newer = d.resolve("後加工工賃明細（2026年7月度)V02.xlsx");
        Files.write(newer, new byte[] {'P', 'K'});
        Files.setLastModifiedTime(newer, FileTime.fromMillis(System.currentTimeMillis() + 10_000));
        Files.write(d.resolve("後加工工賃明細（2026年8月度)V01.xlsx"), new byte[] {'P', 'K'});
        Files.write(d.resolve("~$後加工工賃明細（2026年8月度)V01.xlsx"), new byte[] {'P', 'K'});
        Map<YearMonthKey, Path> m = TrendScan.listKokubuFiles(d);
        assertEquals(2, m.size());
        assertTrue(m.containsKey(new YearMonthKey(2026, 7)));
        assertTrue(m.containsKey(new YearMonthKey(2026, 8)));
        assertTrue(m.get(new YearMonthKey(2026, 7)).getFileName().toString().endsWith("V02.xlsx"));
    }

    @Test
    @DisplayName("複数年度フォルダを走査し、同月は更新日時が新しい方を採用する")
    void listKokubuMergesDirsPicksNewer() throws Exception {
        Path d2025 = tmp.resolve("工賃明細２０２５年度");
        Path d2026 = tmp.resolve("工賃明細2026年度");
        Files.createDirectories(d2025);
        Files.createDirectories(d2026);
        Path older = d2025.resolve("後加工工賃明細（2026年2月度).xlsx");
        Path newer = d2026.resolve("後加工工賃明細（2026年2月度)V02.xlsx");
        Files.write(older, new byte[] {'P', 'K'});
        Files.write(newer, new byte[] {'P', 'K'});
        Files.setLastModifiedTime(older, FileTime.fromMillis(System.currentTimeMillis() - 100_000));
        Files.setLastModifiedTime(newer, FileTime.fromMillis(System.currentTimeMillis() + 10_000));
        Files.write(d2026.resolve("後加工工賃明細（2026年8月度)V01.xlsx"), new byte[] {'P', 'K'});
        Map<YearMonthKey, Path> m = TrendScan.listKokubuFiles(List.of(d2025, d2026));
        assertEquals(newer, m.get(new YearMonthKey(2026, 2)));
        assertTrue(m.containsKey(new YearMonthKey(2026, 8)));
    }

    @Test
    @DisplayName("pickPeriod は最大月から直近 N か月")
    void pickPeriodEndIsMaxYm() {
        Map<YearMonthKey, Path> available =
                Map.of(
                        new YearMonthKey(2026, 6), Path.of("a"),
                        new YearMonthKey(2026, 8), Path.of("b"),
                        new YearMonthKey(2026, 7), Path.of("c"));
        TrendScan.Period period = TrendScan.pickPeriod(available, 3);
        assertEquals(
                List.of(new YearMonthKey(2026, 6), new YearMonthKey(2026, 7), new YearMonthKey(2026, 8)),
                period.months());
        assertEquals(Path.of("b"), period.files().get(new YearMonthKey(2026, 8)));
        assertFalse(period.files().containsKey(new YearMonthKey(2026, 5)));
    }

    @Test
    @DisplayName("湖南はルート直下と yyyy年度試算 を拾い、バックアップは除外")
    void listKonanFilesFromNendoAndRoot() throws Exception {
        Path root = tmp.resolve("2 後加工試算");
        Files.createDirectories(root);
        writeKonanShisan(root.resolve("月度加工賃試算.xlsx"), "2026年9月度");
        Path nendo = root.resolve("2026年度試算　湖南");
        writeKonanShisan(nendo.resolve("8月度加工賃試算.xlsx"), "2026年8月度");
        writeKonanShisan(nendo.resolve("7月度加工賃試算.xlsx"), "2026年7月度");
        Files.createDirectories(root.resolve("バックアップ"));
        writeKonanShisan(root.resolve("バックアップ").resolve("x.xlsx"), "2026年6月度");
        Map<YearMonthKey, Path> m = TrendScan.listKonanFiles(root);
        assertEquals(3, m.size());
        assertTrue(m.containsKey(new YearMonthKey(2026, 7)));
        assertTrue(m.containsKey(new YearMonthKey(2026, 8)));
        assertTrue(m.containsKey(new YearMonthKey(2026, 9)));
        assertFalse(m.containsKey(new YearMonthKey(2026, 6)));
    }

    private static void writeKonanShisan(Path path, String ymLabel) throws IOException {
        Files.createDirectories(path.getParent());
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet ws = wb.createSheet("東レT.V.C");
            ws.createRow(0).createCell(0).setCellValue(ymLabel);
            try (var out = Files.newOutputStream(path)) {
                wb.write(out);
            }
        }
    }
}
