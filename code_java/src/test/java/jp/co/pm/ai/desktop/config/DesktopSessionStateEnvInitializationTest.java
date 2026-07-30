package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

import java.util.List;

import org.junit.jupiter.api.Test;

class DesktopSessionStateEnvInitializationTest {

    private static final ObjectMapper JSON = new ObjectMapper();

    @Test
    void withoutEnvInitializationFields_clearsEnvPathsAndRows() throws Exception {
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
        DesktopSessionState cleared = state.withoutEnvInitializationFields();
        assertEquals("dark", cleared.uiTheme());
        assertEquals("", cleared.excludeRulesPath());
        assertEquals("", cleared.mainRunWorkbook());
        assertEquals("", cleared.mainRunScriptDir());
        assertTrue(cleared.uiEnvRows().isEmpty());
    }

    @Test
    void mergeFactoryScopedFromPreservingEnvInitialization_keepsCurrentWhenFactoryEmpty() throws Exception {
        ObjectNode currentJson =
                (ObjectNode)
                        JSON.readTree(
                                """
                                {
                                  "uiTheme": "dark",
                                  "excludeRulesPath": "C:\\\\current\\\\exclude.json",
                                  "mainRunWorkbook": "C:\\\\current\\\\book.xlsm",
                                  "mainRunScriptDir": "C:\\\\current\\\\py",
                                  "uiEnvRows": [
                                    { "name": "PM_AI_REPO_ROOT", "value": "C:\\\\current", "description": "" }
                                  ]
                                }
                                """);
        DesktopSessionState current = DesktopSessionStateStore.desktopSessionFromStoredJson(currentJson);
        DesktopSessionState factory = DesktopSessionState.empty();
        DesktopSessionState merged = current.mergeFactoryScopedFromPreservingEnvInitialization(factory);
        assertEquals("C:\\current\\exclude.json", merged.excludeRulesPath());
        assertEquals("C:\\current\\book.xlsm", merged.mainRunWorkbook());
        assertEquals("C:\\current\\py", merged.mainRunScriptDir());
        assertEquals(1, merged.uiEnvRows().size());
        assertEquals("PM_AI_REPO_ROOT", merged.uiEnvRows().get(0).name());
        assertEquals("C:\\current", merged.uiEnvRows().get(0).value());
    }

    @Test
    void mergeFactoryScopedFromPreservingEnvInitialization_takesFactoryWhenNonBlank() throws Exception {
        ObjectNode currentJson =
                (ObjectNode)
                        JSON.readTree(
                                """
                                {
                                  "uiTheme": "dark",
                                  "excludeRulesPath": "C:\\\\current\\\\exclude.json",
                                  "mainRunWorkbook": "C:\\\\current\\\\book.xlsm",
                                  "uiEnvRows": [
                                    { "name": "PM_AI_REPO_ROOT", "value": "C:\\\\current", "description": "" }
                                  ]
                                }
                                """);
        ObjectNode factoryJson =
                (ObjectNode)
                        JSON.readTree(
                                """
                                {
                                  "uiTheme": "light",
                                  "excludeRulesPath": "C:\\\\factory\\\\exclude.json",
                                  "mainRunWorkbook": "C:\\\\factory\\\\book.xlsm",
                                  "mainRunScriptDir": "C:\\\\factory\\\\py",
                                  "uiEnvRows": [
                                    { "name": "PM_AI_REPO_ROOT", "value": "C:\\\\factory", "description": "" }
                                  ]
                                }
                                """);
        DesktopSessionState current = DesktopSessionStateStore.desktopSessionFromStoredJson(currentJson);
        DesktopSessionState factory = DesktopSessionStateStore.desktopSessionFromStoredJson(factoryJson);
        DesktopSessionState merged = current.mergeFactoryScopedFromPreservingEnvInitialization(factory);
        assertEquals("dark", merged.uiTheme());
        assertEquals("C:\\factory\\exclude.json", merged.excludeRulesPath());
        assertEquals("C:\\factory\\book.xlsm", merged.mainRunWorkbook());
        assertEquals("C:\\factory\\py", merged.mainRunScriptDir());
        assertEquals(1, merged.uiEnvRows().size());
        assertEquals("C:\\factory", merged.uiEnvRows().get(0).value());
    }

    @Test
    void mergeFactoryScopedFromPreservingEnvInitialization_prefersFactoryUiEnvRowsWhenPresent() throws Exception {
        ObjectNode currentJson =
                (ObjectNode)
                        JSON.readTree(
                                """
                                {
                                  "uiTheme": "dark",
                                  "uiEnvRows": [
                                    { "name": "PM_AI_REPO_ROOT", "value": "C:\\\\current", "description": "" }
                                  ]
                                }
                                """);
        ObjectNode factoryJson =
                (ObjectNode)
                        JSON.readTree(
                                """
                                {
                                  "uiEnvRows": [
                                    { "name": "PM_AI_REPO_ROOT", "value": "C:\\\\factory", "description": "" }
                                  ]
                                }
                                """);
        DesktopSessionState current = DesktopSessionStateStore.desktopSessionFromStoredJson(currentJson);
        DesktopSessionState factory = DesktopSessionStateStore.desktopSessionFromStoredJson(factoryJson);
        List<UiEnvRowSnapshot> rows =
                current.mergeFactoryScopedFromPreservingEnvInitialization(factory).uiEnvRows();
        assertEquals(1, rows.size());
        assertEquals("C:\\factory", rows.get(0).value());
    }
}
