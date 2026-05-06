package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.math.BigDecimal;
import java.nio.file.Path;
import java.util.Optional;

import org.junit.jupiter.api.Test;

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
}
