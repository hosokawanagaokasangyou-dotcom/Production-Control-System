package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class DesktopSessionStateStoreFactoryResetTest {

    @TempDir
    Path tmpHome;

    @TempDir
    Path tmpRepo;

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
    void buildFactoryResetSession_usesExplicitFactorySiteFileEvenWhenStoreDiffers() throws Exception {
        Path initDir = tmpRepo.resolve("init_setting");
        Files.createDirectories(initDir);
        Files.writeString(
                initDir.resolve(InitSettingPaths.sessionDefaultsFileForFactory(FactorySite.KONAN)),
                "{\"uiTheme\":\"light\"}\n",
                StandardCharsets.UTF_8);
        Files.writeString(
                initDir.resolve(InitSettingPaths.sessionDefaultsFileForFactory(FactorySite.KOKUBU)),
                "{\"uiTheme\":\"dark\"}\n",
                StandardCharsets.UTF_8);
        GlobalInitSettingTarget.save(FactorySite.KONAN);

        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, tmpRepo.toString());
        DesktopSessionState forKokubu =
                DesktopSessionStateStore.buildFactoryResetSession(
                        DesktopSessionState.empty(), ui, FactorySite.KOKUBU);
        assertEquals("dark", forKokubu.uiTheme());

        DesktopSessionState forStored =
                DesktopSessionStateStore.buildFactoryResetSession(DesktopSessionState.empty(), ui);
        assertEquals("light", forStored.uiTheme());
    }
}
