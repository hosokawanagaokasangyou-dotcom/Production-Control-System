package jp.co.pm.ai.planning.stage2.parity;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class Stage2ProductionPlanJsonParityTest {

    @Test
    void compare_identicalTrees(@TempDir Path root) throws Exception {
        Path a = root.resolve("a.json");
        Path b = root.resolve("b.json");
        String json = "{\"format_version\":2,\"sheets\":[]}";
        Files.writeString(a, json, StandardCharsets.UTF_8);
        Files.writeString(b, json, StandardCharsets.UTF_8);
        Stage2ProductionPlanJsonParity.CompareResult r = Stage2ProductionPlanJsonParity.compareFiles(a, b);
        assertTrue(r.identical());
    }

    @Test
    void compare_differentVersion(@TempDir Path root) throws Exception {
        Path a = root.resolve("a.json");
        Path b = root.resolve("b.json");
        Files.writeString(a, "{\"format_version\":2}", StandardCharsets.UTF_8);
        Files.writeString(b, "{\"format_version\":3}", StandardCharsets.UTF_8);
        Stage2ProductionPlanJsonParity.CompareResult r = Stage2ProductionPlanJsonParity.compareFiles(a, b);
        assertFalse(r.identical());
    }
}
