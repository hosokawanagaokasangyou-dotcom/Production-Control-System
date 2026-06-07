package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Map;
import java.util.Optional;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.config.AppPaths;

class RdpCompanionLauncherTest {

    @Test
    void resolveRemoteProgramPath_readsUiThenEmpty() {
        assertEquals(
                Optional.of("C:\\Windows\\System32\\notepad.exe"),
                RdpCompanionLauncher.resolveRemoteProgramPath(
                        Map.of(
                                AppPaths.KEY_PM_AI_RDP_COMPANION_PROGRAM,
                                "C:\\Windows\\System32\\notepad.exe")));
        assertFalse(RdpCompanionLauncher.resolveRemoteProgramPath(Map.of()).isPresent());
    }

    @Test
    void resolveRemoteProgramArgs_readsUi() {
        assertEquals(
                "test.txt",
                RdpCompanionLauncher.resolveRemoteProgramArgs(
                        Map.of(AppPaths.KEY_PM_AI_RDP_COMPANION_PROGRAM_ARGS, "test.txt")));
    }

    @Test
    void formatEmbeddedSummary_includesArgsWhenPresent() {
        assertEquals(
                Optional.of("C:\\Windows\\System32\\notepad.exe foo"),
                RdpCompanionLauncher.formatEmbeddedSummary(
                        Map.of(
                                AppPaths.KEY_PM_AI_RDP_COMPANION_PROGRAM,
                                "C:\\Windows\\System32\\notepad.exe",
                                AppPaths.KEY_PM_AI_RDP_COMPANION_PROGRAM_ARGS,
                                "foo")));
    }
}
