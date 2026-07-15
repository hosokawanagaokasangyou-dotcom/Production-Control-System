package jp.co.pm.ai.planning.stage2.source;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Map;

import org.junit.jupiter.api.Test;

class Stage2SourceGuardSnapshotTest {

    @Test
    void matches_rejectsTodayDispatchOffToOn() {
        var started = snapshot(false, false, false, false, Map.of("plan", "a.xlsx"));
        var current = snapshot(true, false, false, false, Map.of("plan", "a.xlsx"));

        assertFalse(started.matches(current));
    }

    @Test
    void matches_rejectsTodayDispatchOnToOff() {
        var started = snapshot(true, false, false, false, Map.of("plan", "a.xlsx"));
        var current = snapshot(false, false, false, false, Map.of("plan", "a.xlsx"));

        assertFalse(started.matches(current));
    }

    @Test
    void matches_rejectsPlanOrDispatchDirtyChange() {
        var started = snapshot(true, false, false, false, Map.of("plan", "a.xlsx"));

        assertFalse(started.matches(snapshot(true, true, false, false, Map.of("plan", "a.xlsx"))));
        assertFalse(started.matches(snapshot(true, false, true, false, Map.of("plan", "a.xlsx"))));
    }

    @Test
    void matches_rejectsAnotherStageStartingDuringGuard() {
        var started = snapshot(true, false, false, false, Map.of("plan", "a.xlsx"));
        var current = snapshot(true, false, false, true, Map.of("plan", "a.xlsx"));

        assertFalse(started.matches(current));
    }

    @Test
    void matches_rejectsDirtyGenerationChangeEvenWhenValueReturnsToOriginal() {
        var started =
                new Stage2SourceGuardSnapshot(
                        true, false, 10L, false, 20L, false, Map.of("plan", "a.xlsx"));
        var current =
                new Stage2SourceGuardSnapshot(
                        true, false, 12L, false, 22L, false, Map.of("plan", "a.xlsx"));

        assertFalse(started.matches(current));
    }

    @Test
    void matches_rejectsEnvironmentChangeWithoutExposingValues() {
        var started = snapshot(true, false, false, false, Map.of("credential", "secret-a"));
        var current = snapshot(true, false, false, false, Map.of("credential", "secret-b"));

        assertFalse(started.matches(current));
        assertFalse(started.mismatchMessage(current).contains("secret-a"));
        assertFalse(started.mismatchMessage(current).contains("secret-b"));
    }

    @Test
    void matches_acceptsUnchangedExecutionState() {
        var started = snapshot(true, false, false, false, Map.of("plan", "a.xlsx"));

        assertTrue(started.matches(snapshot(true, false, false, false, Map.of("plan", "a.xlsx"))));
    }

    private static Stage2SourceGuardSnapshot snapshot(
            boolean todayDispatch,
            boolean planDirty,
            boolean dispatchDirty,
            boolean pipelineRunning,
            Map<String, String> environment) {
        return new Stage2SourceGuardSnapshot(
                todayDispatch,
                planDirty,
                0L,
                dispatchDirty,
                0L,
                pipelineRunning,
                environment);
    }
}
