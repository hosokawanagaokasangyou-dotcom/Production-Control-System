package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class Stage35BaselineActualSnapshotStoreTest {

    @TempDir Path tempDir;

    @Test
    void writeWithMeta_persistsStage35AppliedEvenWhenEntriesEmpty() throws Exception {
        Path dispatchJson = tempDir.resolve("結果_配台表.json");
        Files.writeString(dispatchJson, "{}", StandardCharsets.UTF_8);
        Path overrides = tempDir.resolve("overtime_simulation_overrides.json");
        Files.writeString(
                overrides,
                """
                {
                  "format_version": 1,
                  "working_overrides": { "2026-05-28": { "山田": true } },
                  "overtime_minutes": {}
                }
                """,
                StandardCharsets.UTF_8);

        Stage35BaselineActualSnapshotStore.writeWithMeta(
                dispatchJson,
                Map.of(),
                overrides,
                new Stage35BaselineActualSnapshotStore.OverrideSummary(1, 0, 0));

        Stage35BaselineActualSnapshotStore.Stage35TrialMeta meta =
                Stage35BaselineActualSnapshotStore.tryLoadMeta(dispatchJson);
        assertTrue(meta.hasTrialApplied());
        assertTrue(meta.stage35Applied());
        assertEquals(1, meta.overrideSummary().workOn());
        assertTrue(meta.entries().isEmpty());
        assertTrue(
                Files.isRegularFile(Stage35BaselineActualSnapshotStore.sidecarPathFor(dispatchJson)));
    }

    @Test
    void tryLoadMeta_legacyEntriesOnlySidecar() throws Exception {
        Path dispatchJson = tempDir.resolve("結果_配台表.json");
        Files.writeString(dispatchJson, "{}", StandardCharsets.UTF_8);
        Path sidecar = Stage35BaselineActualSnapshotStore.sidecarPathFor(dispatchJson);
        Files.writeString(
                sidecar,
                """
                {
                  "source_json": "%s",
                  "entries": { "A\\u0001M\\u00012026-05-28": 100.0 }
                }
                """
                        .formatted(dispatchJson.toAbsolutePath().normalize()),
                StandardCharsets.UTF_8);

        Stage35BaselineActualSnapshotStore.Stage35TrialMeta meta =
                Stage35BaselineActualSnapshotStore.tryLoadMeta(dispatchJson);
        assertTrue(meta.hasTrialApplied());
        assertEquals(100.0, meta.entries().values().iterator().next(), 1e-6);
    }

    @Test
    void tryLoadMeta_rejectsMismatchedSourceJson() throws Exception {
        Path dispatchJson = tempDir.resolve("結果_配台表.json");
        Files.writeString(dispatchJson, "{}", StandardCharsets.UTF_8);
        Path sidecar = Stage35BaselineActualSnapshotStore.sidecarPathFor(dispatchJson);
        Files.writeString(
                sidecar,
                """
                {
                  "source_json": "/other/path.json",
                  "stage35_applied": true,
                  "entries": {}
                }
                """,
                StandardCharsets.UTF_8);

        assertFalse(
                Stage35BaselineActualSnapshotStore.tryLoadMeta(dispatchJson).hasTrialApplied());
    }
}
