package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.FileTime;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class Stage2OutputNamingParityTest {

    @Test
    void newestPrimaryPlanJsonAfter_findsNewerFile(@TempDir Path root) throws Exception {
        Path dir = root.resolve("out");
        Files.createDirectories(dir);
        Path old = dir.resolve("計画2401011200001111.json");
        Path neu = dir.resolve("計画2401011200002222.json");
        writeJson(old, "{}");
        long floor = Files.getLastModifiedTime(old).toMillis();
        writeJson(neu, "{\"a\":1}");
        Files.setLastModifiedTime(neu, FileTime.fromMillis(floor + 1));
        Path found = Stage2OutputNaming.newestPrimaryPlanJsonAfter(dir, floor);
        assertNotNull(found);
    }

    @Test
    void newestPrimaryPlanJsonAfter_noneWhenNothingNewer(@TempDir Path root) throws Exception {
        Path dir = root.resolve("out");
        Files.createDirectories(dir);
        Path only = dir.resolve("計画2401011200001111.json");
        writeJson(only, "{}");
        long floor = Stage2OutputNaming.maxPrimaryPlanJsonLastModifiedMillis(dir);
        assertNull(Stage2OutputNaming.newestPrimaryPlanJsonAfter(dir, floor));
    }

    private static void writeJson(Path path, String body) throws Exception {
        Files.writeString(path, body);
    }
}
