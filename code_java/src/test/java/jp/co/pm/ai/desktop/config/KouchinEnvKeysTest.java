package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.Test;

class KouchinEnvKeysTest {

    @Test
    void yearDirsIsNotFolderPickerKey() {
        assertFalse(AppPaths.isFolderPathEnvKey(AppPaths.KEY_PM_AI_KOUCHIN_KOKUBU_YEAR_DIRS));
        assertTrue(AppPaths.isFolderPathEnvKey(AppPaths.KEY_PM_AI_KOUCHIN_BASE_DIR));
        assertTrue(AppPaths.isFolderPathEnvKey(AppPaths.KEY_PM_AI_KOUCHIN_TORAY_CSV_DIR));
        assertTrue(AppPaths.isFolderPathEnvKey(AppPaths.KEY_PM_AI_KOUCHIN_KONAN_TORAY_CSV_DIR));
        assertTrue(AppPaths.isKouchinEnvKey(AppPaths.KEY_PM_AI_KOUCHIN_BASE_DIR));
        assertTrue(AppPaths.isKouchinEnvKey(AppPaths.KEY_PM_AI_KOUCHIN_KOKUBU_YEAR_DIRS));
        assertFalse(AppPaths.isKouchinEnvKey("PM_AI_OUTPUT_DIR"));
    }

    @Test
    void uiRefContainsAllKouchinKeysAndEnvVarDocs() throws Exception {
        String json;
        try (var in = KouchinEnvKeysTest.class.getResourceAsStream(
                "/jp/co/pm/ai/desktop/ui_ref_env_defaults.json")) {
            assertTrue(in != null);
            json = new String(in.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8);
        }
        for (String key : List.of(
                AppPaths.KEY_PM_AI_KOUCHIN_BASE_DIR,
                AppPaths.KEY_PM_AI_KOUCHIN_TORAY_CSV_DIR,
                AppPaths.KEY_PM_AI_KOUCHIN_KONAN_TORAY_CSV_DIR,
                AppPaths.KEY_PM_AI_KOUCHIN_KOKUBU_NAGAOKA_DIR,
                AppPaths.KEY_PM_AI_KOUCHIN_KOKUBU_ALADDIN_DIR,
                AppPaths.KEY_PM_AI_KOUCHIN_KONAN_SHISAN_DIR,
                AppPaths.KEY_PM_AI_KOUCHIN_KONAN_ALADDIN_DIR,
                AppPaths.KEY_PM_AI_KOUCHIN_KONAN_MONTHLY_DIR,
                AppPaths.KEY_PM_AI_KOUCHIN_KOKUBU_YEAR_DIRS,
                AppPaths.KEY_PM_AI_KOUCHIN_OUTPUT_DIR,
                AppPaths.KEY_PM_AI_KOUCHIN_JUDGMENT_DIR)) {
            assertTrue(json.contains("\"" + key + "\""), key);
            assertTrue(EnvVarDocs.logicOnly(key) != null && !EnvVarDocs.logicOnly(key).isBlank(), key);
        }
    }

    @Test
    void kokubuNagaokaDefaultIsYearFolderNotAutoVerifyCopy() throws Exception {
        assertTrue(AppPaths.DEFAULT_KOUCHIN_KOKUBU_NAGAOKA_DIR.contains("工賃明細2026年度"));
        assertFalse(AppPaths.DEFAULT_KOUCHIN_KOKUBU_NAGAOKA_DIR.contains("●自動検証"));
        String json;
        try (var in = KouchinEnvKeysTest.class.getResourceAsStream(
                "/jp/co/pm/ai/desktop/ui_ref_env_defaults.json")) {
            assertTrue(in != null);
            json = new String(in.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8);
        }
        int keyAt = json.indexOf("\"PM_AI_KOUCHIN_KOKUBU_NAGAOKA_DIR\"");
        assertTrue(keyAt >= 0);
        String around = json.substring(keyAt, Math.min(json.length(), keyAt + 400));
        assertTrue(around.contains("工賃明細2026年度"), around);
    }
}
