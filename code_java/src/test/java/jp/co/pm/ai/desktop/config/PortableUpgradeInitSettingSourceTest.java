package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class PortableUpgradeInitSettingSourceTest {

    @Test
    void factoryReset_usesExplicitPortableInitSettingDirectory(@TempDir Path tmp) throws Exception {
        Path oldRepoInit = tmp.resolve("old-repo").resolve("init_setting");
        Path portableInit = tmp.resolve("install").resolve("pm-ai-data").resolve("init_setting");
        Files.createDirectories(oldRepoInit);
        Files.createDirectories(portableInit);
        Files.writeString(
                oldRepoInit.resolve(InitSettingPaths.sessionDefaultsFileForFactory(FactorySite.KOKUBU)),
                "{\"uiTheme\":\"stale\"}\n",
                StandardCharsets.UTF_8);
        Files.writeString(
                portableInit.resolve(InitSettingPaths.sessionDefaultsFileForFactory(FactorySite.KOKUBU)),
                "{\"uiTheme\":\"portable\"}\n",
                StandardCharsets.UTF_8);

        DesktopSessionState state =
                DesktopSessionStateStore.buildFactoryResetSessionFromInitSettingOnly(
                        portableInit, FactorySite.KOKUBU);

        assertEquals("portable", state.uiTheme());
        assertTrue(Files.exists(portableInit.resolve("session_defaults_kokubu.json")));
    }

    @Test
    void overwrite_usesExplicitDestinationRoot(@TempDir Path tmp) throws Exception {
        Path sourceRoot = tmp.resolve("install").resolve("pm-ai-data");
        Path destinationRoot = tmp.resolve("current-repo");
        Path sourceInit = sourceRoot.resolve("init_setting");
        Files.createDirectories(sourceInit);
        Files.writeString(
                sourceInit.resolve(InitSettingPaths.SESSION_DEFAULTS_FILE),
                "{\"uiTheme\":\"portable\"}\n",
                StandardCharsets.UTF_8);

        InitSettingPersistence.applyPortableUpgradeOverwriteFromPmAiData(
                sourceRoot, destinationRoot);

        assertEquals(
                "{\"uiTheme\":\"portable\"}\n",
                Files.readString(
                        destinationRoot
                                .resolve("init_setting")
                                .resolve(InitSettingPaths.SESSION_DEFAULTS_FILE),
                        StandardCharsets.UTF_8));
    }
}
