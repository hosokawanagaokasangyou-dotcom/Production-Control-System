package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class PipelineStageGateTest {

    private static PipelineStageGate.State state(boolean planInput, boolean stage2) {
        return new PipelineStageGate.State(planInput, stage2);
    }

    @Test
    void stage20NeedsPlanInput() {
        assertFalse(PipelineStageGate.canRunStage20(state(false, false)));
        assertTrue(PipelineStageGate.canRunStage20(state(true, false)));
    }

    @Test
    void stage21NeedsPlanInputLikeStage20() {
        var noPlan = state(false, false);
        var planOnly = state(true, false);
        var both = state(true, true);
        assertFalse(PipelineStageGate.canRunStage21(noPlan));
        assertTrue(PipelineStageGate.canRunStage21(planOnly));
        assertTrue(PipelineStageGate.canRunStage21(both));
        assertFalse(PipelineStageGate.stage21DisabledReason(noPlan).isEmpty());
        assertTrue(PipelineStageGate.stage21DisabledReason(planOnly).isEmpty());
        assertTrue(PipelineStageGate.stage21DisabledReason(both).isEmpty());
    }

    @Test
    void nullStateIsNeverRunnable() {
        assertFalse(PipelineStageGate.canRunStage20(null));
        assertFalse(PipelineStageGate.canRunStage21(null));
    }
}
