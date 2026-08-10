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

class GlobalInitSettingTargetTest {

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
    void peekEffective_doesNotPersistInferredFactory() throws Exception {
        Files.createDirectories(tmpHome.resolve(".pm-ai-desktop"));
        Files.writeString(
                tmpHome.resolve(".pm-ai-desktop/global-init-setting-target-factory.txt"),
                "KONAN",
                StandardCharsets.UTF_8);
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                        FactorySite.KOKUBU.taskInputSourceDir());
        assertEquals(FactorySite.KOKUBU, GlobalInitSettingTarget.peekEffective(ui));
        assertEquals(FactorySite.KONAN, GlobalInitSettingTarget.load());
    }

    @Test
    void loadEffective_syncsStoreFromUiEnv() throws Exception {
        Files.createDirectories(tmpHome.resolve(".pm-ai-desktop"));
        Files.writeString(
                tmpHome.resolve(".pm-ai-desktop/global-init-setting-target-factory.txt"),
                "KONAN",
                StandardCharsets.UTF_8);
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                        FactorySite.KOKUBU.taskInputSourceDir());
        assertEquals(FactorySite.KOKUBU, GlobalInitSettingTarget.loadEffective(ui));
        assertEquals(FactorySite.KOKUBU, GlobalInitSettingTarget.load());
    }

    @Test
    void loadEffective_fallsBackToStoredValueWhenUiAmbiguous() throws Exception {
        Files.createDirectories(tmpHome.resolve(".pm-ai-desktop"));
        Files.writeString(
                tmpHome.resolve(".pm-ai-desktop/global-init-setting-target-factory.txt"),
                "KOKUBU",
                StandardCharsets.UTF_8);
        assertEquals(FactorySite.KOKUBU, GlobalInitSettingTarget.loadEffective(Map.of()));
    }

    @Test
    void loadEffective_keepsStoredKokubuWhenEnvSignalsTie() throws Exception {
        Files.createDirectories(tmpHome.resolve(".pm-ai-desktop"));
        Files.writeString(
                tmpHome.resolve(".pm-ai-desktop/global-init-setting-target-factory.txt"),
                "KOKUBU",
                StandardCharsets.UTF_8);
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR,
                        FactorySite.KONAN.portableBundleSourceDir(),
                        AppPaths.KEY_PM_AI_MASTER_WORKBOOK,
                        FactorySite.KOKUBU.pmAiMasterWorkbookEnvValue(Map.of()));
        assertEquals(FactorySite.KOKUBU, GlobalInitSettingTarget.loadEffective(ui));
        assertEquals(FactorySite.KOKUBU, GlobalInitSettingTarget.load());
    }
}
