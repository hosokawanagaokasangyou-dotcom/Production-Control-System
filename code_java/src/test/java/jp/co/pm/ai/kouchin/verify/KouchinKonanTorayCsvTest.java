package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;

class KouchinKonanTorayCsvTest {

    @TempDir
    Path tmp;

    @Test
    @DisplayName("湖南①の既定は湖南共有DATAの東レCSV")
    void konanDefaultIsSharedDataTorayCsv() {
        assertEquals(
                "\\\\192.168.0.101\\共有フォルダ\\湖南工場\\湖南共有\\002  加工G\\●配台AIシステム\\共有DATA\\東レCSV",
                AppPaths.DEFAULT_KOUCHIN_KONAN_TORAY_CSV_DIR);
        assertEquals("PM_AI_KOUCHIN_KONAN_TORAY_CSV_DIR", AppPaths.KEY_PM_AI_KOUCHIN_KONAN_TORAY_CSV_DIR);
        assertTrue(AppPaths.isFolderPathEnvKey(AppPaths.KEY_PM_AI_KOUCHIN_KONAN_TORAY_CSV_DIR));
        assertTrue(AppPaths.isKouchinEnvKey(AppPaths.KEY_PM_AI_KOUCHIN_KONAN_TORAY_CSV_DIR));
    }

    @Test
    @DisplayName("取り込み先は現工場の①フォルダ")
    void importDirFollowsCurrentFactory() {
        Path kokubu = tmp.resolve("kokubu-csv");
        Path konan = tmp.resolve("konan-csv");
        KouchinPaths konanSite = KouchinPaths.fromEnv(Map.of(
                AppPaths.KEY_PM_AI_KOUCHIN_TORAY_CSV_DIR, kokubu.toString(),
                AppPaths.KEY_PM_AI_KOUCHIN_KONAN_TORAY_CSV_DIR, konan.toString(),
                AppPaths.KEY_PM_AI_FACTORY_SITE, "KONAN"));
        assertEquals(konan.toAbsolutePath().normalize(), konanSite.importTorayCsvDir().toAbsolutePath().normalize());
        KouchinPaths kokubuSite = KouchinPaths.fromEnv(Map.of(
                AppPaths.KEY_PM_AI_KOUCHIN_TORAY_CSV_DIR, kokubu.toString(),
                AppPaths.KEY_PM_AI_KOUCHIN_KONAN_TORAY_CSV_DIR, konan.toString(),
                AppPaths.KEY_PM_AI_FACTORY_SITE, "KOKUBU"));
        assertEquals(kokubu.toAbsolutePath().normalize(), kokubuSite.importTorayCsvDir().toAbsolutePath().normalize());
    }

    @Test
    @DisplayName("検証の①は現工場フォルダを優先し、無ければ他工場")
    void verifyPrefersCurrentFactoryCsvThenFallback() throws Exception {
        Path kokubu = tmp.resolve("kokubu-csv");
        Path konan = tmp.resolve("konan-csv");
        Files.createDirectories(kokubu);
        Files.createDirectories(konan);
        Files.writeString(kokubu.resolve("RVSHEET202608.csv"), "k");
        Files.writeString(konan.resolve("RVSHEET202609.csv"), "n");
        KouchinPaths konanSite = KouchinPaths.fromEnv(Map.of(
                AppPaths.KEY_PM_AI_KOUCHIN_TORAY_CSV_DIR, kokubu.toString(),
                AppPaths.KEY_PM_AI_KOUCHIN_KONAN_TORAY_CSV_DIR, konan.toString(),
                AppPaths.KEY_PM_AI_FACTORY_SITE, "KONAN"));
        assertEquals(konan.toAbsolutePath().normalize(), konanSite.resolveTorayCsvDir().toAbsolutePath().normalize());
        Files.delete(konan.resolve("RVSHEET202609.csv"));
        assertEquals(kokubu.toAbsolutePath().normalize(), konanSite.resolveTorayCsvDir().toAbsolutePath().normalize());
    }

    @Test
    @DisplayName("検出は工場ごとの①フォルダを見る")
    void scanUsesFactoryOwnTorayDir() throws Exception {
        Path kokubu = tmp.resolve("kokubu-csv");
        Path konan = tmp.resolve("konan-csv");
        Files.createDirectories(kokubu);
        Files.createDirectories(konan);
        Files.writeString(kokubu.resolve("RVSHEET202608.csv"), "k");
        Files.writeString(konan.resolve("RVSHEET202609.csv"), "n");
        KouchinPaths paths = KouchinPaths.fromEnv(Map.of(
                AppPaths.KEY_PM_AI_KOUCHIN_TORAY_CSV_DIR, kokubu.toString(),
                AppPaths.KEY_PM_AI_KOUCHIN_KONAN_TORAY_CSV_DIR, konan.toString(),
                AppPaths.KEY_PM_AI_FACTORY_SITE, "KONAN"));
        assertTrue(KouchinDiscovery.scan(FactoryId.KOKUBU, paths).stream()
                .anyMatch(r -> r.fullPath().endsWith("RVSHEET202608.csv")));
        assertTrue(KouchinDiscovery.scan(FactoryId.KONAN, paths).stream()
                .anyMatch(r -> r.fullPath().endsWith("RVSHEET202609.csv")));
    }
}
