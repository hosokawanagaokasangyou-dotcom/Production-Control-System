package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class FileDiscoveryUiScanTest {

    @TempDir
    Path tmp;

    @Test
    @DisplayName("年度フォルダ+「m月度加工賃試算」からシートを開かずに年月を取る")
    void guessesYmFromNendoDirAndMonthFileName() {
        Path file = tmp.resolve("2026年度試算　湖南").resolve("7月度加工賃試算.xlsm");
        assertEquals(Optional.of(new YearMonthKey(2026, 7)), FileDiscovery.guessShisanYmFromPath(file));
    }

    @Test
    @DisplayName("ファイル名に「yyyy年m月度」があればそれを使う")
    void guessesYmFromGatsudoFileName() {
        Path file = tmp.resolve("2026年7月度加工賃試算.xlsm");
        assertEquals(Optional.of(new YearMonthKey(2026, 7)), FileDiscovery.guessShisanYmFromPath(file));
    }

    @Test
    @DisplayName("作業中の「月度加工賃試算.xlsm」はパスだけでは年月不明")
    void workingShisanFileHasNoYmInPath() {
        assertTrue(FileDiscovery.guessShisanYmFromPath(tmp.resolve("月度加工賃試算.xlsm")).isEmpty());
    }

    @Test
    @DisplayName("湖南②の当月検出はダミーxlsmでもファイル名だけで当月を選ぶ")
    void findsCurrentShisanByFileNameWithoutOpeningWorkbook() throws Exception {
        Path root = tmp.resolve("shisan");
        Path yearDir = root.resolve("2026年度試算　湖南");
        Files.createDirectories(yearDir);
        Path june = yearDir.resolve("6月度加工賃試算.xlsm");
        Path july = yearDir.resolve("7月度加工賃試算.xlsm");
        Files.writeString(june, "not-excel");
        Files.writeString(july, "not-excel");
        Files.writeString(root.resolve("月度加工賃試算.xlsm"), "not-excel");

        YearMonthKey ym = new YearMonthKey(2026, 7);
        MonthlyFileSet set = FileDiscovery.findShisanFiles(root, ym);
        assertEquals(july, set.current());
    }

    @Test
    @DisplayName("③は新しいファイルから開き、対象月が一致したら残りを開かない")
    void findAladdinStopsAtNewestMatchingYm() throws Exception {
        Path dir = tmp.resolve("aladdin");
        Files.createDirectories(dir);
        Path olderWrong = dir.resolve("依頼NO別問合せ_20260101_000000.xlsx");
        Files.writeString(olderWrong, "not-excel");
        Path newestMatch = dir.resolve("依頼NO別問合せ_20260731_120000.xlsx");
        writeAladdin(newestMatch, "対象年月 : 2026年07月");

        Path found = FileDiscovery.findAladdin(dir, new YearMonthKey(2026, 7));
        assertEquals(newestMatch, found);
    }

    @Test
    @DisplayName("UI検出はダミー試算でも①②③の行を返す")
    void scanDoesNotRequireValidShisanWorkbook() throws Exception {
        Path csvDir = tmp.resolve("csv");
        Path shisanDir = tmp.resolve("shisan");
        Path aladdinDir = tmp.resolve("aladdin");
        Path yearDir = shisanDir.resolve("2026年度試算　湖南");
        Files.createDirectories(csvDir);
        Files.createDirectories(yearDir);
        Files.createDirectories(aladdinDir);
        Files.write(
                csvDir.resolve("RVSHEET202607.csv"),
                "入庫場所\n".getBytes(Charset.forName("windows-31j")));
        Path july = yearDir.resolve("7月度加工賃試算.xlsm");
        Files.writeString(july, "not-excel");
        Path aladdin = aladdinDir.resolve("依頼NO別問合せ_20260731_120000.xlsx");
        writeAladdin(aladdin, "対象年月 : 2026年07月");

        KouchinPaths paths = new KouchinPaths(
                csvDir,
                tmp.resolve("nagaoka"),
                tmp.resolve("kokubu3"),
                shisanDir,
                aladdinDir,
                tmp.resolve("monthly"),
                tmp.resolve("out"),
                tmp.resolve("judge"));
        List<KouchinDiscovery.Row> rows = KouchinDiscovery.scan(FactoryId.KONAN, paths);
        assertEquals("2026年度試算　湖南/7月度加工賃試算.xlsm", row(rows, "②加工賃試算").path());
        assertEquals(july.toString(), row(rows, "②加工賃試算").fullPath());
        assertTrue(row(rows, "③アラジン").path().endsWith("依頼NO別問合せ_20260731_120000.xlsx"), row(rows, "③アラジン").path());
        assertTrue(row(rows, "①東レCSV").path().endsWith("RVSHEET202607.csv"), row(rows, "①東レCSV").path());
        assertEquals(false, row(rows, "②加工賃試算").missing());
        assertEquals(false, row(rows, "③アラジン").missing());
    }

    @Test
    @DisplayName("UIパスは参照フォルダからの相対で、区切りは /")
    void uiDisplayPathUsesRelativeSlash() {
        Path root = tmp.resolve("shisan");
        Path file = root.resolve("2026年度試算　湖南").resolve("7月度加工賃試算.xlsm");
        assertEquals("2026年度試算　湖南/7月度加工賃試算.xlsm", FileDiscovery.uiDisplayPath(root, file));
        assertEquals("csv/RVSHEET202608.csv", FileDiscovery.uiDisplayPath(tmp.resolve("csv"), tmp.resolve("csv").resolve("RVSHEET202608.csv")));
    }

    @Test
    @DisplayName("見つからないときは末尾2段のフォルダ名を / で出す")
    void uiDisplayFolderKeepsLastTwoSegments() {
        Path dir = tmp.resolve("湖南工場").resolve("2 後加工試算");
        assertEquals("湖南工場/2 後加工試算", FileDiscovery.uiDisplayFolder(dir));
    }

    private static KouchinDiscovery.Row row(List<KouchinDiscovery.Row> rows, String role) {
        return rows.stream().filter(r -> role.equals(r.role())).findFirst().orElseThrow();
    }

    private static void writeAladdin(Path path, String taisho) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet ws = wb.createSheet("問合せ");
            ws.createRow(0).createCell(0).setCellValue(taisho);
            Row h = ws.createRow(1);
            h.createCell(0).setCellValue("依頼NO");
            h.createCell(1).setCellValue("項目");
            h.createCell(2).setCellValue("--合計--");
            try (var os = Files.newOutputStream(path)) {
                wb.write(os);
            }
        }
    }
}
