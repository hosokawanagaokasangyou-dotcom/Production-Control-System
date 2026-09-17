package jp.co.pm.ai.desktop.kouchin;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.kouchin.verify.KouchinDiscovery;
import jp.co.pm.ai.kouchin.verify.KouchinPaths;

class KouchinUnusedSource3Test {

    @TempDir
    Path tmp;

    @Test
    @DisplayName("現工場の③があるとき他工場の③は不使用")
    void dimOtherFactorySource3WhenCurrentExists() throws Exception {
        Path kokubuDir = tmp.resolve("kokubu-aladdin");
        Path konanDir = tmp.resolve("konan-jisseki");
        Files.createDirectories(kokubuDir);
        Files.createDirectories(konanDir);
        Path kFile = kokubuDir.resolve("依頼NO別問合せ_k.xlsx");
        Path nFile = konanDir.resolve("依頼NO別問合せ_n.xlsx");
        Files.write(kFile, new byte[] {1});
        Files.write(nFile, new byte[] {1});
        KouchinPaths paths = KouchinPaths.fromEnv(Map.of(
                AppPaths.KEY_PM_AI_KOUCHIN_KOKUBU_ALADDIN_DIR, kokubuDir.toString(),
                AppPaths.KEY_PM_AI_KOUCHIN_KONAN_ALADDIN_DIR, konanDir.toString()));
        var k3 = new KouchinDiscovery.Row(
                KouchinDiscovery.ROLE_3, kFile.getFileName().toString(), kFile.toString(), "2026年8月度", false, "");
        var n3 = new KouchinDiscovery.Row(
                KouchinDiscovery.ROLE_3, nFile.getFileName().toString(), nFile.toString(), "2026年8月度", false, "");
        var lines = KouchinVerifyTabController.markUnusedSource3(
                KouchinVerifyTabController.buildDiscoveryLines(List.of(k3), List.of(n3)),
                FactorySite.KONAN,
                paths);
        assertEquals(2, lines.size());
        assertTrue(find(lines, kFile).isUnused());
        assertFalse(find(lines, nFile).isUnused());
        assertTrue(find(lines, kFile).getNote().contains("現工場以外"));
        assertEquals("pm-kouchin-unused-row", find(lines, kFile).getRowCss());
        assertEquals("", find(lines, nFile).getRowCss());

        var kokubuFirst = KouchinVerifyTabController.markUnusedSource3(
                KouchinVerifyTabController.buildDiscoveryLines(List.of(k3), List.of(n3)),
                FactorySite.KOKUBU,
                paths);
        assertFalse(find(kokubuFirst, kFile).isUnused());
        assertTrue(find(kokubuFirst, nFile).isUnused());
    }

    @Test
    @DisplayName("現工場の①があるとき他工場の①は不使用")
    void dimOtherFactoryTorayCsvWhenCurrentExists() throws Exception {
        Path kokubuDir = tmp.resolve("kokubu-csv");
        Path konanDir = tmp.resolve("konan-csv");
        Files.createDirectories(kokubuDir);
        Files.createDirectories(konanDir);
        Path kFile = kokubuDir.resolve("RVSHEET202608.csv");
        Path nFile = konanDir.resolve("RVSHEET202609.csv");
        Files.writeString(kFile, "k");
        Files.writeString(nFile, "n");
        KouchinPaths paths = KouchinPaths.fromEnv(Map.of(
                AppPaths.KEY_PM_AI_KOUCHIN_TORAY_CSV_DIR, kokubuDir.toString(),
                AppPaths.KEY_PM_AI_KOUCHIN_KONAN_TORAY_CSV_DIR, konanDir.toString()));
        var k1 = new KouchinDiscovery.Row(
                KouchinDiscovery.ROLE_1, kFile.getFileName().toString(), kFile.toString(), "2026年8月度", false, "");
        var n1 = new KouchinDiscovery.Row(
                KouchinDiscovery.ROLE_1, nFile.getFileName().toString(), nFile.toString(), "2026年9月度", false, "");
        var lines = KouchinVerifyTabController.markUnusedSource3(
                KouchinVerifyTabController.buildDiscoveryLines(List.of(k1), List.of(n1)),
                FactorySite.KONAN,
                paths);
        assertTrue(find(lines, kFile).isUnused());
        assertFalse(find(lines, nFile).isUnused());
        assertEquals("国分/湖南共通", find(lines, nFile).getFactory());

        var kokubuFirst = KouchinVerifyTabController.markUnusedSource3(
                KouchinVerifyTabController.buildDiscoveryLines(List.of(k1), List.of(n1)),
                FactorySite.KOKUBU,
                paths);
        assertFalse(find(kokubuFirst, kFile).isUnused());
        assertTrue(find(kokubuFirst, nFile).isUnused());
    }

    @Test
    @DisplayName("現工場の③が無いときは他工場の③を暗転しない")
    void doNotDimSoleSource3FromOtherFactory() throws Exception {
        Path kokubuDir = tmp.resolve("kokubu-aladdin");
        Path konanDir = tmp.resolve("konan-jisseki");
        Files.createDirectories(kokubuDir);
        Files.createDirectories(konanDir);
        Path nFile = konanDir.resolve("依頼NO別問合せ_n.xlsx");
        Files.write(nFile, new byte[] {1});
        KouchinPaths paths = KouchinPaths.fromEnv(Map.of(
                AppPaths.KEY_PM_AI_KOUCHIN_KOKUBU_ALADDIN_DIR, kokubuDir.toString(),
                AppPaths.KEY_PM_AI_KOUCHIN_KONAN_ALADDIN_DIR, konanDir.toString()));
        var missing = new KouchinDiscovery.Row(
                KouchinDiscovery.ROLE_3, kokubuDir.toString(), kokubuDir.toString(), "", true, "見つかりません");
        var n3 = new KouchinDiscovery.Row(
                KouchinDiscovery.ROLE_3, nFile.getFileName().toString(), nFile.toString(), "2026年8月度", false, "");
        var lines = KouchinVerifyTabController.markUnusedSource3(
                KouchinVerifyTabController.buildDiscoveryLines(List.of(missing), List.of(n3)),
                FactorySite.KOKUBU,
                paths);
        assertFalse(find(lines, nFile).isUnused());
        assertFalse(findMissing(lines).isUnused());
    }

    @Test
    @DisplayName("同一パスの③は暗転しない")
    void doNotDimWhenBothFactoriesShareSameSource3() throws Exception {
        Path sharedDir = tmp.resolve("shared");
        Files.createDirectories(sharedDir);
        Path file = sharedDir.resolve("依頼NO別問合せ.xlsx");
        Files.write(file, new byte[] {1});
        KouchinPaths paths = KouchinPaths.fromEnv(Map.of(
                AppPaths.KEY_PM_AI_KOUCHIN_KOKUBU_ALADDIN_DIR, sharedDir.toString(),
                AppPaths.KEY_PM_AI_KOUCHIN_KONAN_ALADDIN_DIR, sharedDir.toString()));
        var row = new KouchinDiscovery.Row(
                KouchinDiscovery.ROLE_3, file.getFileName().toString(), file.toString(), "2026年8月度", false, "");
        var lines = KouchinVerifyTabController.markUnusedSource3(
                KouchinVerifyTabController.buildDiscoveryLines(List.of(row), List.of(row)),
                FactorySite.KONAN,
                paths);
        assertFalse(lines.get(0).isUnused());
        assertFalse(lines.get(1).isUnused());
    }

    @Test
    @DisplayName("不使用行は表行を暗転する")
    void unusedRowStyleIsWired() throws Exception {
        String src = Files.readString(Path.of(
                "src/main/java/jp/co/pm/ai/desktop/kouchin/KouchinVerifyTabController.java"));
        assertTrue(src.contains("pm-kouchin-unused-row"), src);
        assertTrue(src.contains("markUnusedSource3("), src);
        String css = Files.readString(Path.of("src/main/resources/jp/co/pm/ai/desktop/css/pm-ai-desktop.css"));
        assertTrue(css.contains(".table-row-cell.pm-kouchin-unused-row"), css);
        assertTrue(css.contains("-fx-opacity"), css);
    }

    private static KouchinVerifyTabController.DiscoveryLine find(
            List<KouchinVerifyTabController.DiscoveryLine> lines, Path file) {
        String want = file.toString();
        return lines.stream()
                .filter(l -> l.source() != null && want.equals(l.source().fullPath()))
                .findFirst()
                .orElseThrow();
    }

    private static KouchinVerifyTabController.DiscoveryLine findMissing(
            List<KouchinVerifyTabController.DiscoveryLine> lines) {
        return lines.stream().filter(KouchinVerifyTabController.DiscoveryLine::isMissing).findFirst().orElseThrow();
    }
}
