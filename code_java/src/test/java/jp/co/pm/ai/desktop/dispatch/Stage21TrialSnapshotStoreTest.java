package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class Stage21TrialSnapshotStoreTest {

    @TempDir Path tempDir;

    @Test
    void writePromotedMeta_persistsPromotedFlagWithoutComparisonBaseline() throws Exception {
        Path dispatchJson = tempDir.resolve("結果_配台表.json");
        Files.writeString(dispatchJson, "{\"columns\":[],\"rows\":[]}");
        Path overrides = tempDir.resolve("overtime_simulation_overrides.json");
        Files.writeString(overrides, "{\"working_overrides\":{\"work_on\":[]}}");

        Stage21TrialSnapshotStore.writePromotedMeta(
                dispatchJson,
                overrides,
                new Stage21TrialSnapshotStore.OverrideSummary(2, 1, 3));

        Stage21TrialSnapshotStore.Stage21TrialMeta meta =
                Stage21TrialSnapshotStore.tryLoadMeta(dispatchJson);
        assertTrue(meta.hasPromotedToMain());
        assertTrue(meta.hasAttendanceMeta());
        assertFalse(meta.hasComparisonBaseline());
        assertFalse(meta.hasTrialApplied() && meta.hasComparisonBaseline());
    }

    @Test
    void writeWithMeta_persistsStage21AppliedEvenWhenEntriesEmpty() throws Exception {
        Path dispatchJson = tempDir.resolve("結果_配台表.json");
        Files.writeString(dispatchJson, "{\"columns\":[],\"rows\":[]}");
        Path stage21Json = tempDir.resolve("stage21").resolve("結果_配台表.json");
        Files.createDirectories(stage21Json.getParent());
        Files.writeString(stage21Json, "{}");
        Path overrides = tempDir.resolve("stage21").resolve("overtime_simulation_overrides.json");
        Files.writeString(overrides, "{\"working_overrides\":{\"work_on\":[]}}");

        Stage21TrialSnapshotStore.writeWithMeta(
                dispatchJson,
                Map.of(),
                stage21Json,
                overrides,
                new Stage21TrialSnapshotStore.OverrideSummary(1, 0, 0));

        Stage21TrialSnapshotStore.Stage21TrialMeta meta =
                Stage21TrialSnapshotStore.tryLoadMeta(dispatchJson);
        assertTrue(meta.hasTrialApplied());
        assertTrue(meta.stage21Applied());
        assertEquals(stage21Json.toAbsolutePath().normalize().toString(), meta.stage21ResultDispatchJson());
        assertTrue(
                Files.isRegularFile(Stage21TrialSnapshotStore.sidecarPathFor(dispatchJson)));
    }

    @Test
    void tryLoadMeta_rejectsMismatchedSourceJson() throws Exception {
        Path dispatchJson = tempDir.resolve("結果_配台表.json");
        Files.writeString(dispatchJson, "{}");
        Path sidecar = Stage21TrialSnapshotStore.sidecarPathFor(dispatchJson);
        Files.createDirectories(sidecar.getParent());
        Files.writeString(
                sidecar,
                """
                {
                  "source_json": "/other/path/結果_配台表.json",
                  "stage21_applied": true,
                  "entries": {}
                }
                """);
        assertFalse(Stage21TrialSnapshotStore.tryLoadMeta(dispatchJson).hasTrialApplied());
    }
}
