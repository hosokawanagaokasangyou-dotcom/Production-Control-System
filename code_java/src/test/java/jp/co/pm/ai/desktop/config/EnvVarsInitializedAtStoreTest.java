package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class EnvVarsInitializedAtStoreTest {

    @TempDir
    Path tmpHome;

    private String previousUserHome;

    @BeforeEach
    void isolateStore() {
        previousUserHome = System.getProperty("user.home");
        System.setProperty("user.home", tmpHome.toString());
    }

    @AfterEach
    void restoreUserHome() {
        if (previousUserHome != null) {
            System.setProperty("user.home", previousUserHome);
        } else {
            System.clearProperty("user.home");
        }
    }

    @Test
    void load_emptyWhenMissing() {
        assertFalse(EnvVarsInitializedAtStore.load().isPresent());
        assertFalse(EnvVarsInitializedAtStore.isRecorded());
        assertEquals("—", EnvVarsInitializedAtStore.formatForToolbar());
    }

    @Test
    void recordNow_persistsAndFormats() throws Exception {
        Instant fixed = Instant.parse("2026-07-30T07:44:00Z");
        Files.createDirectories(tmpHome.resolve(".pm-ai-desktop"));
        Files.writeString(
                EnvVarsInitializedAtStore.storePathForTests(),
                fixed.toString(),
                StandardCharsets.UTF_8);
        assertEquals(fixed, EnvVarsInitializedAtStore.load().orElseThrow());
        assertTrue(EnvVarsInitializedAtStore.isRecorded());
        assertTrue(EnvVarsInitializedAtStore.formatForToolbar().contains("2026"));
    }
}
