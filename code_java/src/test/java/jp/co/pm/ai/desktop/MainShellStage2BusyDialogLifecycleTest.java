package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class MainShellStage2BusyDialogLifecycleTest {

    @Test
    void shouldCloseStageRunBusyForPostStage2AsyncWork_trueForLongRunningFollowUps() {
        assertTrue(
                MainShellController.shouldCloseStageRunBusyForPostStage2AsyncWork(
                        MainRunStage2Progress.State.DISPATCH_RELOADING));
        assertTrue(
                MainShellController.shouldCloseStageRunBusyForPostStage2AsyncWork(
                        MainRunStage2Progress.State.DELIVERY_RELOADING));
        assertTrue(
                MainShellController.shouldCloseStageRunBusyForPostStage2AsyncWork(
                        MainRunStage2Progress.State.EXCEL_GENERATING));
    }

    @Test
    void shouldCloseStageRunBusyForPostStage2AsyncWork_falseWhilePythonRuns() {
        assertFalse(
                MainShellController.shouldCloseStageRunBusyForPostStage2AsyncWork(
                        MainRunStage2Progress.State.RUNNING));
    }
}
