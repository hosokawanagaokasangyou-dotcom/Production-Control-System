package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

class EnvVarsInitialTemplateTest {

  @Test
  void matches_returnsTrueWhenCurrentEqualsExpected() {
    Map<String, String> expected = Map.of("PM_AI_REPO_ROOT", "C:\\repo", "PM_AI_OUTPUT_DIR", "");
    Map<String, String> current =
        Map.of("PM_AI_REPO_ROOT", "C:\\repo", "PM_AI_OUTPUT_DIR", "", "EXTRA", "x");
    assertTrue(EnvVarsInitialTemplate.matches(current, expected, k -> true));
  }

  @Test
  void matches_returnsFalseWhenValueDiffers() {
    Map<String, String> expected = Map.of("PM_AI_REPO_ROOT", "C:\\repo");
    Map<String, String> current = Map.of("PM_AI_REPO_ROOT", "C:\\other");
    assertFalse(EnvVarsInitialTemplate.matches(current, expected, k -> true));
  }

  @Test
  void overlayFactorySiteValues_setsFactoryManagedKeys() {
    LinkedHashMap<String, String> map = new LinkedHashMap<>();
    map.put(AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR, "old");
    map.put(AppPaths.KEY_PM_AI_FACTORY_SITE, "old");
    EnvVarsInitialTemplate.overlayFactorySiteValues(map, FactorySite.KOKUBU, Map.of());
    assertTrue(map.get(AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR).contains("国分"));
    assertEquals("KOKUBU", map.get(AppPaths.KEY_PM_AI_FACTORY_SITE));
  }

  @Test
  void buildExpectedMap_appliesBootstrapDefaultsForEmptyBootstrapKeys() {
    Map<String, String> built =
        EnvVarsInitialTemplate.buildExpectedMap(
            List.of(AppPaths.KEY_PM_AI_REPO_ROOT, AppPaths.KEY_PM_AI_SKIP_WORKBOOK_ENV_SHEET),
            FactorySite.KONAN,
            (k, ui) ->
                AppPaths.KEY_PM_AI_SKIP_WORKBOOK_ENV_SHEET.equals(k) ? "1" : "bootstrap-" + k,
            (k, ui) -> "",
            k -> true);
    assertTrue(built.containsKey(AppPaths.KEY_PM_AI_REPO_ROOT));
    assertEquals("1", built.get(AppPaths.KEY_PM_AI_SKIP_WORKBOOK_ENV_SHEET));
  }
}
