package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.util.Map;

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

    @Test
    void envFingerprint_roundTrip() {
        Map<String, String> env =
                Map.of(
                        "PM_AI_REPO_ROOT",
                        "C:\\repo",
                        "PM_AI_OUTPUT_DIR",
                        "",
                        "B_KEY",
                        "2",
                        "A_KEY",
                        "1");
        EnvVarsInitializedAtStore.recordEnvFingerprint(env, k -> k.startsWith("PM_AI_"));
        assertTrue(EnvVarsInitializedAtStore.envFingerprintMatches(env, k -> k.startsWith("PM_AI_")));
        assertFalse(
                EnvVarsInitializedAtStore.envFingerprintMatches(
                        Map.of("PM_AI_REPO_ROOT", "C:\\other"), k -> k.startsWith("PM_AI_")));
    }

    @Test
    void matchesRecordedBaselineForKeys_ignoresRdpDrift() throws Exception {
        Map<String, String> baseline =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        "C:\\repo",
                        AppPaths.KEY_PM_AI_REQUEST_FORM_RDP_PROFILE,
                        "");
        EnvVarsInitializedAtStore.recordEnvFingerprint(baseline, k -> true);
        Map<String, String> current =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        "C:\\repo",
                        AppPaths.KEY_PM_AI_REQUEST_FORM_RDP_PROFILE,
                        "C:\\signed.rdp");
        assertTrue(
                EnvVarsInitializedAtStore.matchesRecordedBaselineForKeys(
                        current,
                        k -> !RemoteDesktopEnvRows.excludedFromMainShellEnvInitFingerprint(k)));
    }

    @Test
    void envFingerprint_ignoresPipelineRuntimeSyncedKeys() {
        Map<String, String> baseline =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        "C:\\repo",
                        AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON,
                        "");
        EnvVarsInitializedAtStore.recordEnvFingerprint(
                baseline, k -> !AppPaths.isPipelineRuntimeSyncedEnvKey(k));
        Map<String, String> afterStageRun =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        "C:\\repo",
                        AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON,
                        "C:\\output\\stage1_exclude_rules.json");
        assertTrue(
                EnvVarsInitializedAtStore.envFingerprintMatches(
                        afterStageRun, k -> !AppPaths.isPipelineRuntimeSyncedEnvKey(k)));
    }

    @Test
    void envFingerprint_ignoresRdpTabKeys() {
        Map<String, String> baseline =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        "C:\\repo",
                        AppPaths.KEY_PM_AI_REQUEST_FORM_RDP_PROFILE,
                        "");
        EnvVarsInitializedAtStore.recordEnvFingerprint(
                baseline,
                k -> !RemoteDesktopEnvRows.excludedFromMainShellEnvInitFingerprint(k));
        Map<String, String> afterRdpVisit =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        "C:\\repo",
                        AppPaths.KEY_PM_AI_REQUEST_FORM_RDP_PROFILE,
                        "C:\\profiles\\signed.rdp");
        assertTrue(
                EnvVarsInitializedAtStore.envFingerprintMatches(
                        afterRdpVisit,
                        k ->
                                !RemoteDesktopEnvRows.excludedFromMainShellEnvInitFingerprint(k)));
    }
}
