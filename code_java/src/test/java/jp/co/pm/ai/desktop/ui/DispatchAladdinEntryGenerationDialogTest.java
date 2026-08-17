package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;

class DispatchAladdinEntryGenerationDialogTest {

    @TempDir
    Path tempDir;

    @Test
    void generationRoot_usesAladdinEntryDispatchPlanDir() {
        Path repo = tempDir.resolve("repo");
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, repo.toString());

        assertEquals(
                AppPaths.aladdinEntryDispatchPlanDir(ui),
                DispatchAladdinEntryGenerationDialog.generationRoot(ui));
    }

    @Test
    void listCandidatePaths_prependsLocalLatestWhenRequested() throws Exception {
        Path repo = tempDir.resolve("repo");
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, repo.toString());
        Path local = AppPaths.aladdinEntryDispatchPlanLocalXlsxPath(ui);
        Files.createDirectories(local.getParent());
        Files.writeString(local, "local");
        Path genDir = DispatchAladdinEntryGenerationDialog.generationRoot(ui).resolve("tester");
        Files.createDirectories(genDir);
        Path generation = genDir.resolve("アラジン入力用_配台計画_20260817-120000.xlsx");
        Files.writeString(generation, "gen");

        List<Path> withLocal =
                DispatchAladdinEntryGenerationDialog.listCandidatePaths(ui, "tester", true);
        List<Path> generationsOnly =
                DispatchAladdinEntryGenerationDialog.listCandidatePaths(ui, "tester", false);

        assertEquals(local.toAbsolutePath().normalize(), withLocal.get(0).toAbsolutePath().normalize());
        assertTrue(withLocal.stream().anyMatch(p -> p.getFileName().equals(generation.getFileName())));
        assertEquals(1, generationsOnly.size());
        assertEquals(generation.toAbsolutePath().normalize(), generationsOnly.get(0).toAbsolutePath().normalize());
    }
}
