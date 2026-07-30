package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class DesktopSessionStateStoreFactoryResetTest {

    private static final ObjectMapper JSON = new ObjectMapper();

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

    private void writeSessionState(String json) throws Exception {
        Path store = AppPaths.resolveSessionStateStorePath();
        Files.createDirectories(store.getParent());
        Files.writeString(store, json, StandardCharsets.UTF_8);
    }

    private ObjectNode readSessionStateRoot() throws Exception {
        return (ObjectNode) JSON.readTree(AppPaths.resolveSessionStateStorePath().toFile());
    }

    @Test
    void applyPortableUpgradeBundledPolicyToSessionStore_skipsEnvKeysFromBundledMerge() throws Exception {
        writeSessionState(
                """
                {
                  "uiTheme": "dark",
                  "mainRunWorkbook": "C:\\\\user\\\\book.xlsm",
                  "mainRunScriptDir": "C:\\\\user\\\\py",
                  "excludeRulesPath": "C:\\\\user\\\\exclude.json",
                  "uiEnvRows": [
                    { "name": "PM_AI_REPO_ROOT", "value": "C:\\\\user", "description": "" }
                  ]
                }
                """);
        Path initDir = tmpRepo.resolve("init_setting");
        Files.createDirectories(initDir);
        Files.writeString(
                initDir.resolve(InitSettingPaths.sessionDefaultsFileForFactory(FactorySite.KONAN)),
                """
                {
                  "uiTheme": "light",
                  "mainRunWorkbook": "C:\\\\bundled\\\\book.xlsm",
                  "mainRunScriptDir": "C:\\\\bundled\\\\py",
                  "excludeRulesPath": "C:\\\\bundled\\\\exclude.json",
                  "uiEnvRows": [
                    { "name": "PM_AI_REPO_ROOT", "value": "C:\\\\bundled", "description": "" }
                  ]
                }
                """,
                StandardCharsets.UTF_8);
        GlobalInitSettingTarget.save(FactorySite.KONAN);
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, tmpRepo.toString());

        DesktopSessionStateStore.applyPortableUpgradeBundledPolicyToSessionStore(ui);

        ObjectNode root = readSessionStateRoot();
        assertEquals("light", root.get("uiTheme").asText());
        assertEquals("C:\\user\\book.xlsm", root.get("mainRunWorkbook").asText());
        assertEquals("C:\\user\\py", root.get("mainRunScriptDir").asText());
        assertEquals("C:\\user\\exclude.json", root.get("excludeRulesPath").asText());
        assertEquals("C:\\user", root.get("uiEnvRows").get(0).get("value").asText());
    }

    @Test
    void toJsonObjectForGlobalInitSetting_omitsEnvInitializationFields() throws Exception {
        ObjectNode in =
                (ObjectNode)
                        JSON.readTree(
                                """
                                {
                                  "uiTheme": "dark",
                                  "excludeRulesPath": "C:\\\\exclude.json",
                                  "mainRunWorkbook": "C:\\\\master.xlsm",
                                  "mainRunScriptDir": "C:\\\\python",
                                  "uiEnvRows": [
                                    { "name": "PM_AI_REPO_ROOT", "value": "C:\\\\repo", "description": "" }
                                  ]
                                }
                                """);
        DesktopSessionState state = DesktopSessionStateStore.desktopSessionFromStoredJson(in);
        ObjectNode root = DesktopSessionStateStore.toJsonObjectForGlobalInitSetting(state);
        assertFalse(root.has("uiEnvRows"));
        assertFalse(root.has("excludeRulesPath"));
        assertFalse(root.has("mainRunWorkbook"));
        assertFalse(root.has("mainRunScriptDir"));
        assertEquals("dark", root.get("uiTheme").asText());
    }

    @Test
    void buildFactoryResetSessionFromInitSettingOnly_ignoresEnvRowsInInitSettingFile() throws Exception {
        Path initDir = tmpRepo.resolve("init_setting");
        Files.createDirectories(initDir);
        Files.writeString(
                initDir.resolve(InitSettingPaths.sessionDefaultsFileForFactory(FactorySite.KONAN)),
                """
                {
                  "uiTheme": "light",
                  "excludeRulesPath": "C:\\\\legacy\\\\exclude.json",
                  "mainRunWorkbook": "C:\\\\legacy\\\\master.xlsm",
                  "uiEnvRows": [ { "name": "PM_AI_REPO_ROOT", "value": "C:\\\\legacy", "description": "" } ]
                }
                """,
                StandardCharsets.UTF_8);
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, tmpRepo.toString());
        DesktopSessionState loaded =
                DesktopSessionStateStore.buildFactoryResetSessionFromInitSettingOnly(
                        ui, FactorySite.KONAN);
        assertEquals("light", loaded.uiTheme());
        assertEquals("", loaded.excludeRulesPath());
        assertEquals("", loaded.mainRunWorkbook());
        assertEquals("", loaded.mainRunScriptDir());
        assertTrue(loaded.uiEnvRows().isEmpty());
    }

    @Test
    void buildFactoryResetSessionFromInitSettingOnly_usesExplicitFactorySiteFileEvenWhenStoreDiffers()
            throws Exception {
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
                DesktopSessionStateStore.buildFactoryResetSessionFromInitSettingOnly(
                        ui, FactorySite.KOKUBU);
        assertEquals("dark", forKokubu.uiTheme());

        DesktopSessionState forStored =
                DesktopSessionStateStore.buildFactoryResetSessionFromInitSettingOnly(ui, FactorySite.KONAN);
        assertEquals("light", forStored.uiTheme());
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
