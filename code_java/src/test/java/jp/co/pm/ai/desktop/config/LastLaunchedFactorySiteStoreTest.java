package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class LastLaunchedFactorySiteStoreTest {

    private String priorHome;
    private String priorUserHome;

    @BeforeEach
    void setUp(@TempDir Path tmp) {
        priorHome = AppPaths.desktopAppHomeDirName();
        priorUserHome = System.getProperty("user.home");
        System.setProperty("user.home", tmp.toString());
        AppPaths.setDesktopAppHomeDirName(".pm-ai-desktop-test");
        LastLaunchedFactorySiteStore.resetForTests();
    }

    @AfterEach
    void tearDown() {
        LastLaunchedFactorySiteStore.resetForTests();
        AppPaths.setDesktopAppHomeDirName(priorHome);
        System.setProperty("user.home", priorUserHome);
    }

    @Test
    void load_whenFileMissing_returnsKonan() {
        assertFalse(Files.isRegularFile(AppPaths.resolveLastLaunchedFactorySiteStorePath()));
        assertEquals(FactorySite.KONAN, LastLaunchedFactorySiteStore.load());
    }

    @Test
    void save_thenLoad_returnsKokubu() throws Exception {
        LastLaunchedFactorySiteStore.save(FactorySite.KOKUBU);
        Path path = AppPaths.resolveLastLaunchedFactorySiteStorePath();
        assertTrue(Files.isRegularFile(path));
        String json = Files.readString(path, StandardCharsets.UTF_8);
        assertTrue(json.contains("\"factorySite\""));
        assertTrue(json.contains("KOKUBU"));
        assertEquals(FactorySite.KOKUBU, LastLaunchedFactorySiteStore.load());
    }

    @Test
    void load_whenJsonInvalid_returnsKonan() throws Exception {
        Path path = AppPaths.resolveLastLaunchedFactorySiteStorePath();
        Files.createDirectories(path.getParent());
        Files.writeString(path, "{not-json", StandardCharsets.UTF_8);
        assertEquals(FactorySite.KONAN, LastLaunchedFactorySiteStore.load());
    }

    @Test
    void load_whenFactorySiteUnknown_returnsKonan() throws Exception {
        Path path = AppPaths.resolveLastLaunchedFactorySiteStorePath();
        Files.createDirectories(path.getParent());
        Files.writeString(path, "{\"factorySite\":\"RDP_LAUNCHER\"}", StandardCharsets.UTF_8);
        assertEquals(FactorySite.KONAN, LastLaunchedFactorySiteStore.load());
    }

    @Test
    void save_ignoresRdpLauncher() {
        LastLaunchedFactorySiteStore.save(FactorySite.RDP_LAUNCHER);
        assertFalse(Files.isRegularFile(AppPaths.resolveLastLaunchedFactorySiteStorePath()));
        assertEquals(FactorySite.KONAN, LastLaunchedFactorySiteStore.load());
    }
}
