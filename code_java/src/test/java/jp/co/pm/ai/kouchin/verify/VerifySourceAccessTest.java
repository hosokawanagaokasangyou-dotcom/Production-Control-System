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
        List<KouchinDiscovery.Row> ok = List.of(found("①東レCSV", csv), found("②長岡明細", x2), found("③アラジン", x3));
        assertTrue(VerifySourceAccess.factorySourcesReady(ok));
        assertFalse(VerifySourceAccess.factorySourcesReady(List.of()));
        assertFalse(VerifySourceAccess.factorySourcesReady(null));

        KouchinDiscovery.Row missing = new KouchinDiscovery.Row(
                "②長岡明細", tmp.toString(), tmp.toString(), "", true, "見つかりません");
        assertFalse(VerifySourceAccess.factorySourcesReady(List.of(found("①東レCSV", csv), missing, found("③アラジン", x3))));
        assertFalse(VerifySourceAccess.factorySourcesReady(
                List.of(found("①東レCSV", csv), found("②長岡明細", tmp.resolve("gone.xlsx")), found("③アラジン", x3))));
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

    private static KouchinDiscovery.Row found(String role, Path file) {
        return new KouchinDiscovery.Row(role, file.getFileName().toString(), file.toString(), "2026年8月度", false, "");
    }
}
