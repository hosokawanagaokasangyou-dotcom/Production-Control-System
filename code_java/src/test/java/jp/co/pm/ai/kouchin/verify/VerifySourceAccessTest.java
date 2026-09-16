package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assertions.assertEquals;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class VerifySourceAccessTest {

    @TempDir
    Path tmp;

    @Test
    @DisplayName("実在ファイルの読取・書込可否を判定する")
    void fileAccessReadAndWrite() throws Exception {
        Path file = tmp.resolve("明細.xlsx");
        Files.writeString(file, "xlsx");
        VerifySourceAccess.FileAccess ok = VerifySourceAccess.FileAccess.of(file);
        assertTrue(ok.present());
        assertTrue(ok.readable());
        assertTrue(ok.writable());
        assertEquals("可", ok.readLabel());
        assertEquals("可", ok.writeLabel());
        assertEquals("pm-kouchin-access-ok", ok.readCss());
        assertEquals("pm-kouchin-access-ok", ok.writeCss());
        assertTrue(VerifySourceAccess.canWriteFile(file));
        assertFalse(VerifySourceAccess.canWriteFile(null));
        assertFalse(VerifySourceAccess.canWriteFile(tmp));
        VerifySourceAccess.FileAccess missing = VerifySourceAccess.FileAccess.of(tmp.resolve("gone.xlsx"));
        assertFalse(missing.present());
        assertEquals("—", missing.readLabel());
        assertEquals("—", missing.writeLabel());
        assertEquals("pm-kouchin-access-na", missing.readCss());
        KouchinDiscovery.Row absent = new KouchinDiscovery.Row(
                "②長岡明細", tmp.toString(), tmp.toString(), "", true, "見つかりません");
        assertEquals("—", VerifySourceAccess.FileAccess.ofRow(absent).writeLabel());
    }

    @Test
    @DisplayName("実在ファイルは読み取り専用でアクセス可")
    void canReadRegularFile() throws Exception {
        Path file = tmp.resolve("RVSHEET202608.csv");
        Files.writeString(file, "csv");
        assertTrue(VerifySourceAccess.canReadAtLeastReadOnly(file));
        assertFalse(VerifySourceAccess.canReadAtLeastReadOnly(null));
        assertFalse(VerifySourceAccess.canReadAtLeastReadOnly(tmp));
        assertFalse(VerifySourceAccess.canReadAtLeastReadOnly(tmp.resolve("missing.csv")));
    }

    @Test
    @DisplayName("検出行がすべて読めるときだけその工場の検証を許可")
    void factoryReadyOnlyWhenEveryDiscoveredFileReadable() throws Exception {
        Path csv = tmp.resolve("RVSHEET202608.csv");
        Path x2 = tmp.resolve("明細.xlsx");
        Path x3 = tmp.resolve("アラジン.xlsx");
        Files.writeString(csv, "a");
        Files.writeString(x2, "b");
        Files.writeString(x3, "c");
        List<KouchinDiscovery.Row> ok = List.of(found("①東レCSV", csv), found("②長岡明細", x2), found("③月次実績", x3));
        assertTrue(VerifySourceAccess.factorySourcesReady(ok));
        assertFalse(VerifySourceAccess.factorySourcesReady(List.of()));
        assertFalse(VerifySourceAccess.factorySourcesReady(null));

        KouchinDiscovery.Row missing = new KouchinDiscovery.Row(
                "②長岡明細", tmp.toString(), tmp.toString(), "", true, "見つかりません");
        assertFalse(VerifySourceAccess.factorySourcesReady(List.of(found("①東レCSV", csv), missing, found("③月次実績", x3))));
        assertFalse(VerifySourceAccess.factorySourcesReady(
                List.of(found("①東レCSV", csv), found("②長岡明細", tmp.resolve("gone.xlsx")), found("③月次実績", x3))));
    }

    @Test
    @DisplayName("まとめて検証は両工場の関連ファイルがすべて読めるときだけ")
    void bothReadyRequiresEachFactory() throws Exception {
        Path csv = tmp.resolve("csv.csv");
        Path k2 = tmp.resolve("k2.xlsx");
        Path k3 = tmp.resolve("k3.xlsx");
        Path n2 = tmp.resolve("n2.xlsm");
        Path n3 = tmp.resolve("n3.xlsx");
        Path monthly = tmp.resolve("monthly.xlsx");
        Files.writeString(csv, "1");
        Files.writeString(k2, "2");
        Files.writeString(k3, "3");
        Files.writeString(n2, "4");
        Files.writeString(n3, "5");
        Files.writeString(monthly, "6");
        List<KouchinDiscovery.Row> kokubu = List.of(found("①", csv), found("②", k2), found("③", k3));
        List<KouchinDiscovery.Row> konan = List.of(
                found("①", csv), found("②", n2), found("③", n3), found("湖南 月次処理", monthly));
        assertTrue(VerifySourceAccess.bothFactoriesReady(kokubu, konan));
        assertFalse(VerifySourceAccess.bothFactoriesReady(kokubu, List.of(found("①", csv))));
        assertEquals("見つかりません: ②長岡明細", VerifySourceAccess.blockReason(List.of(
                found("①東レCSV", csv),
                new KouchinDiscovery.Row("②長岡明細", tmp.toString(), tmp.toString(), "", true, "見つかりません"))));
    }

    @Test
    @DisplayName("湖南の月次処理が無くても①②③が読めれば検証可（Cはスキップ）")
    void monthlyMissingDoesNotBlockKonan() throws Exception {
        Path csv = tmp.resolve("csv.csv");
        Path n2 = tmp.resolve("n2.xlsm");
        Path n3 = tmp.resolve("n3.xlsx");
        Files.writeString(csv, "1");
        Files.writeString(n2, "2");
        Files.writeString(n3, "3");
        KouchinDiscovery.Row monthlyGone = new KouchinDiscovery.Row(
                "湖南 月次処理", tmp.toString(), tmp.toString(), "", true, "見つかりません");
        List<KouchinDiscovery.Row> konan = List.of(
                found("①東レCSV", csv), found("②試算", n2), found("③月次実績", n3), monthlyGone);
        assertTrue(VerifySourceAccess.factorySourcesReady(konan));
        assertEquals(null, VerifySourceAccess.blockReason(konan));
    }

    @Test
    @DisplayName("読取可・書込不可でも検証は許可し、書込は警告色")
    void readableUnwritableStillReadyAndWarnsOnWrite() throws Exception {
        VerifySourceAccess.FileAccess acc = new VerifySourceAccess.FileAccess(true, true, false);
        assertEquals("可", acc.readLabel());
        assertEquals("不可", acc.writeLabel());
        assertEquals("pm-kouchin-access-ok", acc.readCss());
        assertEquals("pm-kouchin-access-warn", acc.writeCss());
        assertTrue(acc.writeHint().contains("検証は読取できれば可"));
        Path csv = tmp.resolve("c.csv");
        Path x2 = tmp.resolve("x2.xlsx");
        Path x3 = tmp.resolve("x3.xlsx");
        Files.writeString(csv, "a");
        Files.writeString(x2, "b");
        Files.writeString(x3, "c");
        List<KouchinDiscovery.Row> rows = List.of(found("①", csv), found("②", x2), found("③", x3));
        assertEquals(null, VerifySourceAccess.blockReason(rows, r -> acc));
        assertTrue(VerifySourceAccess.factorySourcesReady(rows, r -> acc));
    }

    private static KouchinDiscovery.Row found(String role, Path file) {
        return new KouchinDiscovery.Row(role, file.getFileName().toString(), file.toString(), "2026年8月度", false, "");
    }
}
