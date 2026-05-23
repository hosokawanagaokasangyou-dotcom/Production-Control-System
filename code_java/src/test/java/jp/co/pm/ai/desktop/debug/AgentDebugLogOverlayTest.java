package jp.co.pm.ai.desktop.debug;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.Map;
import jp.co.pm.ai.desktop.config.AppPaths;
import org.junit.jupiter.api.Test;

class AgentDebugLogOverlayTest {

    @Test
    void overlayRefreshesDebugLogPathWhenSessionChanges() {
        Map<String, String> env = new LinkedHashMap<>();
        env.put("PM_AI_REPO_ROOT", "C:\\repo");
        env.put("PM_AI_CODE_PYTHON_DIR", "C:\\repo\\code\\python");
        env.put("PM_AI_AGENT_DEBUG_SESSION", "e04a1d");
        AgentDebugLog.overlayPythonChildDebugEnv(env);
        String first = env.get("PM_AI_DEBUG_LOG");
        assertTrue(first != null && first.contains("debug-e04a1d.log"), first);

        env.put("PM_AI_AGENT_DEBUG_SESSION", "a15218");
        AgentDebugLog.overlayPythonChildDebugEnv(env);
        String second = env.get("PM_AI_DEBUG_LOG");
        assertEquals("a15218", env.get("PM_AI_AGENT_DEBUG_SESSION"));
        assertTrue(second != null && second.contains("debug-a15218.log"), second);
    }

    @Test
    void userCursorDebugLogOverridesGeneratedPath() {
        Map<String, String> env = new LinkedHashMap<>();
        env.put("PM_AI_REPO_ROOT", "C:\\repo");
        env.put("PM_AI_CODE_PYTHON_DIR", "C:\\repo\\code\\python");
        env.put(AppPaths.KEY_PM_AI_CURSOR_DEBUG_LOG, "C:\\custom\\my-debug.log");
        env.put("PM_AI_AGENT_DEBUG_SESSION", "a15218");
        AgentDebugLog.overlayPythonChildDebugEnv(env);
        assertEquals(
                Path.of("C:\\custom\\my-debug.log").toAbsolutePath().normalize().toString(),
                env.get("PM_AI_DEBUG_LOG"));
    }
}
