package jp.co.pm.ai.desktop.debug;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.nio.file.Path;
import java.util.Map;
import org.junit.jupiter.api.Test;

class AgentDebugLogCursorRootTest {

    @Test
    void cursorRootUsesParentWhenRepoLeafIsCodeJava() {
        Path repo = Path.of("/tmp/pm-ai-monorepo/code_java");
        assertEquals(
                Path.of("/tmp/pm-ai-monorepo"),
                AgentDebugLog.resolveCursorDebugDirectoryRoot(
                        Map.of("PM_AI_REPO_ROOT", repo.toString())));
    }

    @Test
    void ndjsonPathUnderWorkspaceCursorWhenRepoIsCodeJava() {
        Path repo = Path.of("/tmp/pm-ai-ws/code_java");
        Path resolved =
                AgentDebugLog.resolveNdjsonPath(
                        Map.of("PM_AI_REPO_ROOT", repo.toString()), "a15218");
        assertEquals(
                Path.of("/tmp/pm-ai-ws/.cursor/debug-a15218.log").toAbsolutePath().normalize(),
                resolved.toAbsolutePath().normalize());
    }
}
