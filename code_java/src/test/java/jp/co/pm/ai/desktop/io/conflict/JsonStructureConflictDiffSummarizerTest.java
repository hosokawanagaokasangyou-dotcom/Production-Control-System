package jp.co.pm.ai.desktop.io.conflict;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

class JsonStructureConflictDiffSummarizerTest {

    @Test
    void reportsAddedKeyAndArraySizeChange() {
        Path p = Path.of("rules.json").toAbsolutePath().normalize();
        byte[] base = "{\"rules\":[1],\"meta\":{}}".getBytes(StandardCharsets.UTF_8);
        byte[] disk = "{\"rules\":[1,2,3],\"meta\":{},\"newKey\":true}".getBytes(StandardCharsets.UTF_8);
        String s =
                new JsonStructureConflictDiffSummarizer("配台不要ルール")
                        .summarize(Map.of(p, base), Map.of(p, disk), List.of(p));
        assertTrue(s.contains("newKey"), s);
        assertTrue(s.contains("rules"), s);
        assertTrue(s.contains("配台不要ルール"), s);
    }
}
