package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;

class Stage21OutputPromoterTest {

    @TempDir Path tempDir;

    @Test
    void promoteToMainOutput_copiesDispatchPlanSidecarsAndMember() throws Exception {
        Path mainDir = tempDir.resolve("output");
        Path stage21Dir = mainDir.resolve("stage21");
        Files.createDirectories(stage21Dir);

        String stamp = "2605300712006110";
        Files.writeString(stage21Dir.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME), "{\"ok\":true}");
        Files.writeString(stage21Dir.resolve("overtime_simulation_overrides.json"), "{}");
        Files.writeString(stage21Dir.resolve("計画" + stamp + ".json"), "{}");
        Files.writeString(stage21Dir.resolve("計画" + stamp + "設.json"), "{}");
        Files.writeString(stage21Dir.resolve("計画" + stamp + "表.json"), "{}");
        Files.writeString(stage21Dir.resolve("人員" + stamp + ".json"), "{}");

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_OUTPUT_DIR,
                        mainDir.toString(),
                        AppPaths.KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR,
                        mainDir.toString());

        Stage21OutputPromoter.Result result = Stage21OutputPromoter.promoteToMainOutput(ui);

        assertTrue(result.filesCopied() >= 5);
        assertTrue(Files.isRegularFile(mainDir.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME)));
        assertTrue(Files.isRegularFile(mainDir.resolve("計画" + stamp + ".json")));
        assertTrue(Files.isRegularFile(mainDir.resolve("計画" + stamp + "設.json")));
        assertTrue(Files.isRegularFile(mainDir.resolve("人員" + stamp + ".json")));
        assertEquals(
                mainDir.resolve("計画" + stamp + ".json").toAbsolutePath().normalize(),
                result.mainPlanJson().toAbsolutePath().normalize());
    }

    @Test
    void artifactFamilyPrefix_extractsPlanStampPrefix() {
        Path plan = Path.of("計画2605300712006110.json");
        assertEquals("計画2605300712006110", Stage21OutputPromoter.artifactFamilyPrefix(plan));
        assertTrue(
                Stage21OutputPromoter.belongsToArtifactFamily(
                        Path.of("計画2605300712006110設.json"), "計画2605300712006110"));
        assertFalse(
                Stage21OutputPromoter.belongsToArtifactFamily(
                        Path.of("計画2605300712006999.json"), "計画2605300712006110"));
    }
}
