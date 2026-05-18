package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class PortableBundleSelfUpdaterTest {

    @Test
    void shouldUpdate_whenCanonicalNewer() {
        Optional<BigDecimal> c = Optional.of(new BigDecimal("1.02"));
        Optional<BigDecimal> l = Optional.of(new BigDecimal("1.01"));
        assertTrue(PortableBundleSelfUpdater.shouldUpdate(c, l));
    }

    @Test
    void shouldUpdate_whenLocalMissing_treatAsZero() {
        Optional<BigDecimal> c = Optional.of(new BigDecimal("1.00"));
        Optional<BigDecimal> l = Optional.empty();
        assertTrue(PortableBundleSelfUpdater.shouldUpdate(c, l));
    }

    @Test
    void shouldUpdate_false_whenCanonicalMissing() {
        Optional<BigDecimal> c = Optional.empty();
        Optional<BigDecimal> l = Optional.of(new BigDecimal("2.00"));
        assertFalse(PortableBundleSelfUpdater.shouldUpdate(c, l));
    }

    @Test
    void excludes_output_prefix() {
        assertTrue(PortableBundleSelfUpdater.isExcludedPath(Path.of("output/foo.txt")));
        assertTrue(PortableBundleSelfUpdater.isExcludedPath(Path.of("output")));
    }

    @Test
    void excludes_master_xlsm() {
        assertTrue(PortableBundleSelfUpdater.isExcludedPath(Path.of("master.xlsm")));
        assertFalse(PortableBundleSelfUpdater.isExcludedPath(Path.of("code/other.xlsm")));
    }

    @Test
    void excludes_init_setting_user_profiles_tree() {
        assertTrue(PortableBundleSelfUpdater.isExcludedPath(Path.of("init_setting/user-profiles")));
        assertTrue(PortableBundleSelfUpdater.isExcludedPath(Path.of("init_setting/user-profiles/a.json")));
        assertFalse(PortableBundleSelfUpdater.isExcludedPath(Path.of("init_setting/session_defaults.json")));
        assertFalse(
                PortableBundleSelfUpdater.isExcludedPath(Path.of("init_setting/session_defaults_kokubu.json")));
        assertFalse(
                PortableBundleSelfUpdater.isExcludedPath(Path.of("init_setting/table_column_defaults_konan.json")));
    }

    @Test
    void readLocalBundleVersion_fallsBackToCwdWhenPmAiDataMissing(@TempDir Path tmp) throws IOException {
        Path pm = tmp.resolve("pm-ai-data");
        Files.createDirectories(pm);
        Files.writeString(tmp.resolve(AppPaths.VERSION_TXT_FILE_NAME), "2.01\n", StandardCharsets.UTF_8);
        assertEquals(
                Optional.of(new BigDecimal("2.01")),
                PortableBundleSelfUpdater.readLocalBundleVersion(tmp, pm));
    }

    @Test
    void readCanonicalPortableBundleVersion_readsVersionBesideZip(@TempDir Path tmp) throws IOException {
        Path zip = tmp.resolve("PMD_version_upgrade.zip");
        try (ZipOutputStream zout = new ZipOutputStream(Files.newOutputStream(zip))) {
            zout.putNextEntry(new ZipEntry("pm-ai-data/readme.txt"));
            zout.write(new byte[] {'x'});
            zout.closeEntry();
        }
        Files.writeString(tmp.resolve(AppPaths.VERSION_TXT_FILE_NAME), "9.99\n", StandardCharsets.UTF_8);
        assertEquals(
                Optional.of(new BigDecimal("9.99")),
                PortableBundleSelfUpdater.readCanonicalPortableBundleVersion(zip));
    }

    @Test
    void resolveEffectiveUpgradeZip_fromReleaseFolder(@TempDir Path tmp) throws IOException {
        Path zip = tmp.resolve(PortableBundleSelfUpdater.PORTABLE_UPGRADE_ZIP_NAME);
        Files.writeString(zip, "PK", StandardCharsets.UTF_8);
        assertEquals(Optional.of(zip), PortableBundleSelfUpdater.resolveEffectiveUpgradeZip(tmp));
    }

    @Test
    void readCanonicalPortableBundleVersion_fromReleaseFolder(@TempDir Path tmp) throws IOException {
        Path zip = tmp.resolve(PortableBundleSelfUpdater.PORTABLE_UPGRADE_ZIP_NAME);
        Files.writeString(zip, "PK", StandardCharsets.UTF_8);
        Files.writeString(tmp.resolve(AppPaths.VERSION_TXT_FILE_NAME), "7.50\n", StandardCharsets.UTF_8);
        assertEquals(
                Optional.of(new BigDecimal("7.50")),
                PortableBundleSelfUpdater.readCanonicalPortableBundleVersion(tmp));
    }

    @Test
    void resolveSyncSourceRoot_prefersNestedPmAiData(@TempDir Path tmp) throws IOException {
        Path nested = tmp.resolve("pm-ai-data/code/python");
        Files.createDirectories(nested);
        assertEquals(tmp.resolve("pm-ai-data"), PortableBundleSelfUpdater.resolveSyncSourceRoot(tmp));
    }

    @Test
    void copyUpgradeZipToLocal_copiesBytes(@TempDir Path tmp) throws IOException {
        Path src = tmp.resolve("PMD_version_upgrade.zip");
        Files.writeString(src, "0123456789", StandardCharsets.UTF_8);
        Path local =
                PortableBundleSelfUpdater.copyUpgradeZipToLocal(src, null, null);
        try {
            assertEquals(10, Files.size(local));
        } finally {
            Files.deleteIfExists(local);
        }
    }

    @Test
    void syncFromCanonical_logsEachCopiedRelativePath(@TempDir Path tmp) throws IOException {
        Path canon = tmp.resolve("canon");
        Files.createDirectories(canon.resolve("code"));
        Files.writeString(canon.resolve("code/a.txt"), "a", StandardCharsets.UTF_8);
        Path dest = tmp.resolve("dest");
        List<String> lines = new ArrayList<>();
        PortableBundleSelfUpdater.syncFromCanonical(canon, dest, lines::add);
        assertTrue(lines.stream().anyMatch(s -> s.contains("同期: code/a.txt")));
        assertTrue(lines.stream().anyMatch(s -> s.contains("同期完了: 1 ファイル")));
    }

    @Test
    void shouldUpdateBundle_trueWhenJarHashDiffers(@TempDir Path tmp) throws IOException {
        Path install = tmp.resolve("install");
        Path pm = install.resolve("pm-ai-data");
        Files.createDirectories(pm);
        Files.writeString(install.resolve(AppPaths.VERSION_TXT_FILE_NAME), "1.00\n", StandardCharsets.UTF_8);
        Path app = install.resolve("app");
        Files.createDirectories(app);
        Path jar = app.resolve("production-control-desktop-0.1.0-SNAPSHOT.jar");
        Files.writeString(jar, "same-version", StandardCharsets.UTF_8);
        String hash =
                PortableBundleBuildManifest.sha256Hex(jar).orElseThrow();
        long size = Files.size(jar);

        Path release = tmp.resolve("release");
        Files.createDirectories(release);
        Files.writeString(release.resolve(AppPaths.VERSION_TXT_FILE_NAME), "1.00\n", StandardCharsets.UTF_8);
        String manifest =
                """
                {
                  "version": "1.00",
                  "desktopJarFileName": "production-control-desktop-0.1.0-SNAPSHOT.jar",
                  "desktopJarSize": %d,
                  "desktopJarSha256": "%s"
                }
                """
                        .formatted(size, hash);
        Files.writeString(
                release.resolve(PortableBundleBuildManifest.FILE_NAME),
                manifest,
                StandardCharsets.UTF_8);

        assertFalse(PortableBundleSelfUpdater.shouldUpdateBundle(release, install, pm));
        Files.writeString(jar, "same-version-changed", StandardCharsets.UTF_8);
        assertTrue(PortableBundleSelfUpdater.shouldUpdateBundle(release, install, pm));
    }

    @Test
    void syncDesktopInstallFromBundleRoot_copiesExeAndJar(@TempDir Path tmp) throws IOException {
        Path bundle = tmp.resolve("bundle");
        Files.createDirectories(bundle.resolve("app"));
        Files.writeString(bundle.resolve("PMD.exe"), "exe", StandardCharsets.UTF_8);
        Files.writeString(
                bundle.resolve("app/desktop.jar"), "jar-bytes", StandardCharsets.UTF_8);
        Path install = tmp.resolve("install");
        PortableBundleSelfUpdater.syncDesktopInstallFromBundleRoot(bundle, install, null);
        assertTrue(Files.isRegularFile(install.resolve("PMD.exe")));
        assertEquals("jar-bytes", Files.readString(install.resolve("app/desktop.jar")));
    }

    @Test
    void copyOuterVersionTxtToLocal_writesPmAiDataAndCwd(@TempDir Path tmp) throws IOException {
        Path zip = tmp.resolve(PortableBundleSelfUpdater.PORTABLE_UPGRADE_ZIP_NAME);
        Files.writeString(zip, "PK", StandardCharsets.UTF_8);
        Files.writeString(tmp.resolve(AppPaths.VERSION_TXT_FILE_NAME), "3.14\n", StandardCharsets.UTF_8);
        Path pm = tmp.resolve("install/pm-ai-data");
        Path cwd = tmp.resolve("install");
        Files.createDirectories(pm);
        PortableBundleSelfUpdater.copyOuterVersionTxtToLocal(tmp, cwd, pm);
        assertEquals("3.14", Files.readString(pm.resolve(AppPaths.VERSION_TXT_FILE_NAME)).trim());
        assertEquals("3.14", Files.readString(cwd.resolve(AppPaths.VERSION_TXT_FILE_NAME)).trim());
    }
}
