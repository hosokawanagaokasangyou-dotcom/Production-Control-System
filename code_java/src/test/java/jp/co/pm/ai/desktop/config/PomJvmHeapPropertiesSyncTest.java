package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class PomJvmHeapPropertiesSyncTest {

    @Test
    void formatJvmHeapToken_wholeGiB_usesG() {
        assertEquals("4g", PomJvmHeapPropertiesSync.formatJvmHeapToken(4096));
        assertEquals("8g", PomJvmHeapPropertiesSync.formatJvmHeapToken(8192));
    }

    @Test
    void formatJvmHeapToken_nonWholeGiB_usesM() {
        assertEquals("5121m", PomJvmHeapPropertiesSync.formatJvmHeapToken(5121));
    }

    @Test
    void patchHeapFlagsInText_updatesBatAndCfgLines() {
        String bat =
                "\"%JAVA_EXE%\" -Dfile.encoding=UTF-8 -Xms4g -Xmx4g -XX:+HeapDumpOnOutOfMemoryError";
        assertEquals(
                "\"%JAVA_EXE%\" -Dfile.encoding=UTF-8 -Xms2g -Xmx8g -XX:+HeapDumpOnOutOfMemoryError",
                PomJvmHeapPropertiesSync.patchHeapFlagsInText(bat, "2g", "8g"));

        String cfg =
                "[JavaOptions]\r\n"
                        + "java-options=-Dfile.encoding=UTF-8\r\n"
                        + "java-options=-Xms4g\r\n"
                        + "java-options=-Xmx4g\r\n";
        String patched = PomJvmHeapPropertiesSync.patchHeapFlagsInText(cfg, "2g", "8g");
        assertTrue(patched.contains("java-options=-Xms2g"));
        assertTrue(patched.contains("java-options=-Xmx8g"));
        assertFalse(patched.contains("-Xmx4g"));
    }

    @Test
    void resolveDesktopInstallRoot_findsParentOfPmAiData(@TempDir Path temp) throws Exception {
        Path install = temp.resolve("PMD_install");
        Files.createDirectories(install.resolve("app"));
        Files.createDirectories(install.resolve("runtime"));
        Files.createDirectories(install.resolve("pm-ai-data").resolve("code").resolve("python"));
        Files.writeString(
                install.resolve("pm-ai-data").resolve("code").resolve("python").resolve("task_extract_stage1.py"),
                "# marker",
                StandardCharsets.UTF_8);

        assertTrue(PomJvmHeapPropertiesSync.looksLikeDesktopInstallRoot(install));
    }

    @Test
    void writeJvmHeapFromLaunchPrefs_fixedMode_patchesPortableLauncherBat(@TempDir Path temp) throws Exception {
        Path install = temp.resolve("bundle");
        Files.createDirectories(install.resolve("app"));
        Files.createDirectories(install.resolve("runtime").resolve("bin"));
        Files.createDirectories(install.resolve("pm-ai-data").resolve("code_java"));
        Files.writeString(
                install.resolve("launch-pm-ai-desktop.bat"),
                "java -Xms4g -Xmx4g jp.co.pm.ai.desktop.PmAiFxApp",
                StandardCharsets.UTF_8);
        Files.writeString(
                install.resolve("pm-ai-data").resolve("code_java").resolve("pom.xml"),
                """
                <project>
                  <properties>
                    <jvm.max.heap>4g</jvm.max.heap>
                    <jvm.initial.heap>4g</jvm.initial.heap>
                  </properties>
                </project>
                """,
                StandardCharsets.UTF_8);

        Path prevUserDir = Path.of(System.getProperty("user.dir"));
        System.setProperty("user.dir", install.toString());
        try {
            PomJvmHeapPropertiesSync.writeJvmHeapFromLaunchPrefs(
                    Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, install.resolve("pm-ai-data").toString()),
                    true,
                    6144,
                    8192);
        } finally {
            System.setProperty("user.dir", prevUserDir.toString());
        }

        String bat = Files.readString(install.resolve("launch-pm-ai-desktop.bat"), StandardCharsets.UTF_8);
        assertTrue(bat.contains("-Xmx8g"));
        assertTrue(bat.contains("-Xms6g"));

        String pom =
                Files.readString(
                        install.resolve("pm-ai-data").resolve("code_java").resolve("pom.xml"),
                        StandardCharsets.UTF_8);
        assertTrue(pom.contains("<jvm.max.heap>8g</jvm.max.heap>"));
        assertTrue(pom.contains("<jvm.initial.heap>6g</jvm.initial.heap>"));
    }

    @Test
    void writeJvmHeapFromLaunchPrefs_variableMode_usesMinFloorForInitial(@TempDir Path temp) throws Exception {
        Path install = temp.resolve("bundle");
        Files.createDirectories(install.resolve("pm-ai-data").resolve("code_java"));
        Files.writeString(
                install.resolve("pm-ai-data").resolve("code_java").resolve("pom.xml"),
                """
                <project>
                  <properties>
                    <jvm.max.heap>4g</jvm.max.heap>
                    <jvm.initial.heap>4g</jvm.initial.heap>
                  </properties>
                </project>
                """,
                StandardCharsets.UTF_8);

        Path prevUserDir = Path.of(System.getProperty("user.dir"));
        System.setProperty("user.dir", install.toString());
        try {
            PomJvmHeapPropertiesSync.writeJvmHeapFromLaunchPrefs(
                    Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, install.resolve("pm-ai-data").toString()),
                    false,
                    0,
                    6144);
        } finally {
            System.setProperty("user.dir", prevUserDir.toString());
        }

        String pom =
                Files.readString(
                        install.resolve("pm-ai-data").resolve("code_java").resolve("pom.xml"),
                        StandardCharsets.UTF_8);
        assertTrue(pom.contains("<jvm.max.heap>6g</jvm.max.heap>"));
        assertTrue(pom.contains("<jvm.initial.heap>2g</jvm.initial.heap>"));
    }
}
