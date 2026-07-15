package jp.co.pm.ai.desktop.debug;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.nio.file.Path;
import java.util.Map;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class AgentDebugLogCursorRootTest {

    @Test
    void cursorRootUsesParentWhenRepoLeafIsCodeJava(@TempDir Path tmp) {
        Path repo = tmp.resolve("pm-ai-monorepo").resolve("code_java");
        assertEquals(
                repo.getParent().toAbsolutePath().normalize(),
                AgentDebugLog.resolveCursorDebugDirectoryRoot(
                        Map.of("PM_AI_REPO_ROOT", repo.toString())));
    }

    @Test
    void ndjsonPathUnderWorkspaceCursorWhenRepoIsCodeJava(@TempDir Path tmp) {
        Path repo = tmp.resolve("pm-ai-ws").resolve("code_java");
        Path resolved =
                AgentDebugLog.resolveNdjsonPath(
                        Map.of("PM_AI_REPO_ROOT", repo.toString()), "a15218");
        assertEquals(
                repo.getParent()
                        .resolve(".cursor")
                        .resolve("debug-a15218.log")
                        .toAbsolutePath()
                        .normalize(),
                resolved.toAbsolutePath().normalize());
    }
}
