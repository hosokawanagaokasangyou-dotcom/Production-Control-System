package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.DispatchAladdinEntryWorkbookExporter;

class DispatchAladdinEntryGenerationDialogTest {

    @TempDir
    Path tempDir;

    @Test
    void generationRoot_sharedUsesAladdinEntryDispatchPlanDir() {
        Path repo = tempDir.resolve("repo");
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, repo.toString());

        assertEquals(
                AppPaths.aladdinEntryDispatchPlanDir(ui),
                DispatchAladdinEntryGenerationDialog.generationRoot(
                        ui, DispatchAladdinEntryWorkbookExporter.Destination.SHARED));
    }

    @Test
    void generationRoot_localUsesAladdinEntryDispatchPlanLocalDir() {
        Path repo = tempDir.resolve("repo");
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, repo.toString());

        assertEquals(
                AppPaths.aladdinEntryDispatchPlanLocalDir(ui),
                DispatchAladdinEntryGenerationDialog.generationRoot(
                        ui, DispatchAladdinEntryWorkbookExporter.Destination.LOCAL));
    }
}
