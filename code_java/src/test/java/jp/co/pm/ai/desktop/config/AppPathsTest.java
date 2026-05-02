package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import java.util.Optional;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class AppPathsTest {

    @Test
    void geminiCredentialsJson_usesFilePickerNotFolder() {
        assertTrue(AppPaths.isFilePathEnvKey(AppPaths.KEY_GEMINI_CREDENTIALS_JSON));
        assertFalse(AppPaths.isFolderPathEnvKey(AppPaths.KEY_GEMINI_CREDENTIALS_JSON));
    }

    @Test
    void excludeRulesJson_usesFilePickerNotFolder() {
        assertTrue(AppPaths.isFilePathEnvKey(AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON));
        assertFalse(AppPaths.isFolderPathEnvKey(AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON));
    }

    @Test
    void taskInputSourceDir_defaultMatchesPqAUncSuffix() {
        Path p = AppPaths.resolveTaskInputSourceDir(Map.of());
        String s = p.toString().replace('\\', '/');
        assertTrue(s.contains("192.168.0.101"), "host: " + p);
        assertTrue(s.endsWith("\u751f\u7523\u8a08\u753b\u554f\u5408\u305b"), "suffix: " + p);
    }

    @Test
    void actualDetailSourceDir_defaultMatchesPq02UncSuffix() {
        Path p = AppPaths.resolveActualDetailSourceDir(Map.of());
        String s = p.toString().replace('\\', '/');
        assertTrue(s.contains("192.168.0.101"), "host: " + p);
        assertTrue(s.contains("002"), "segment 002  \u52a0\u5de5G: " + p);
        assertTrue(
                s.endsWith("\u52a0\u5de5\u5b9f\u7e3e\u660e\u7d30DATA"),
                "suffix: " + p);
    }

    @Test
    void resultDispatchTableDir_defaultsToRepoCode(@TempDir Path fakeRepo) throws Exception {
        Path code = fakeRepo.resolve("Production-Control-System").resolve("code").resolve("python");
        Files.createDirectories(code);
        Files.createFile(code.resolve("task_extract_stage1.py"));
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, fakeRepo.resolve("Production-Control-System").toString());
        Path expected =
                fakeRepo.resolve("Production-Control-System").resolve("code").toAbsolutePath().normalize();
        assertEquals(expected, AppPaths.resolveResultDispatchTableDir(ui));
    }

    @Test
    void pickMacroWorkbook_singleFile(@TempDir Path dir) throws Exception {
        Path wb = dir.resolve("only.xlsm");
        Files.createFile(wb);
        assertEquals(Optional.of(wb), AppPaths.pickMacroWorkbook(dir));
    }

    @Test
    void pickMacroWorkbook_prefersNameContainingHaitai(@TempDir Path dir) throws Exception {
        Files.createFile(dir.resolve("other.xlsm"));
        Path preferred = dir.resolve("\u751f\u7523\u7ba1\u7406_AI\u914d\u53f0_V2.xlsm");
        Files.createFile(preferred);
        assertEquals(Optional.of(preferred), AppPaths.pickMacroWorkbook(dir));
    }
}
