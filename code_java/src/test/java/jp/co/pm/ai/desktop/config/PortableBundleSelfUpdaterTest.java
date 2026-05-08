package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
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
        Path zip = tmp.resolve("PMD_version_upgrade_1.zip");
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
}
