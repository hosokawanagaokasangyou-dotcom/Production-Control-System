package jp.co.pm.ai.desktop.runtime;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class WindowsLauncherUserDirTest {

    @Test
    void isPortableBundleInstallRoot_requiresAppAndRuntime(@TempDir Path tmp) throws Exception {
        Files.createDirectories(tmp.resolve("app"));
        Files.createDirectories(tmp.resolve("runtime"));
        assertTrue(WindowsLauncherUserDir.isPortableBundleInstallRoot(tmp));

        Files.createDirectories(tmp.resolve("incomplete").resolve("app"));
        assertFalse(WindowsLauncherUserDir.isPortableBundleInstallRoot(tmp.resolve("incomplete")));
    }
}
