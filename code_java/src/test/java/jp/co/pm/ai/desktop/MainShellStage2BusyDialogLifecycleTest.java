package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class MainShellStage2BusyDialogLifecycleTest {

    @Test
    void shouldKeepStageRunBusyForPostStage2AsyncWork_trueForTabSwitchFollowUps() {
        assertTrue(
                MainShellController.shouldKeepStageRunBusyForPostStage2AsyncWork(
                        MainRunStage2Progress.State.DISPATCH_RELOADING));
        assertTrue(
                MainShellController.shouldKeepStageRunBusyForPostStage2AsyncWork(
                        MainRunStage2Progress.State.DELIVERY_RELOADING));
        assertTrue(
                MainShellController.shouldKeepStageRunBusyForPostStage2AsyncWork(
                        MainRunStage2Progress.State.EXCEL_GENERATING));
    }

    @Test
    void shouldKeepStageRunBusyForPostStage2AsyncWork_falseWhilePythonRuns() {
        assertFalse(
                MainShellController.shouldKeepStageRunBusyForPostStage2AsyncWork(
                        MainRunStage2Progress.State.RUNNING));
    }

    @Test
    void shouldCloseStageRunBusyForPostStage2AsyncWork_falseSoDialogStaysDuringFollowUps() {
        assertFalse(
                MainShellController.shouldCloseStageRunBusyForPostStage2AsyncWork(
                        MainRunStage2Progress.State.DISPATCH_RELOADING));
        assertFalse(
                MainShellController.shouldCloseStageRunBusyForPostStage2AsyncWork(
                        MainRunStage2Progress.State.DELIVERY_RELOADING));
        assertFalse(
                MainShellController.shouldCloseStageRunBusyForPostStage2AsyncWork(
                        MainRunStage2Progress.State.EXCEL_GENERATING));
    }
}
