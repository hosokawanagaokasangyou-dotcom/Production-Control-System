package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class PipelineStageGateTest {

    private static PipelineStageGate.State state(
            boolean planInput, boolean stage2, boolean stage3Input, boolean stage3Result) {
        return new PipelineStageGate.State(planInput, stage2, stage3Input, stage3Result);
    }

    @Test
    void stage20NeedsPlanInput() {
        assertFalse(PipelineStageGate.canRunStage20(state(false, false, false, false)));
        assertTrue(PipelineStageGate.canRunStage20(state(true, false, false, false)));
    }

    @Test
    void stage21AndBuildStage3InputNeedStage2Result() {
        var noResult = state(true, false, false, false);
        var withResult = state(true, true, false, false);
        assertFalse(PipelineStageGate.canRunStage21(noResult));
        assertFalse(PipelineStageGate.canBuildStage3Input(noResult));
        assertTrue(PipelineStageGate.canRunStage21(withResult));
        assertTrue(PipelineStageGate.canBuildStage3Input(withResult));
        assertFalse(PipelineStageGate.stage21DisabledReason(noResult).isEmpty());
        assertTrue(PipelineStageGate.stage21DisabledReason(withResult).isEmpty());
    }

    @Test
    void stage30And32NeedStage3Input() {
        var noInput = state(true, true, false, false);
        var withInput = state(true, true, true, false);
        assertFalse(PipelineStageGate.canRunStage30(noInput));
        assertFalse(PipelineStageGate.canRunStage32(noInput));
        assertTrue(PipelineStageGate.canRunStage30(withInput));
        assertTrue(PipelineStageGate.canRunStage32(withInput));
    }

    @Test
    void stage31NeedsStage3InputAndStage30Result() {
        assertFalse(PipelineStageGate.canRunStage31(state(true, true, true, false)));
        assertTrue(PipelineStageGate.canRunStage31(state(true, true, true, true)));
        assertFalse(
                PipelineStageGate.stage31DisabledReason(state(true, true, false, false)).isEmpty());
        assertTrue(PipelineStageGate.stage31DisabledReason(state(true, true, true, true)).isEmpty());
    }

    @Test
    void nullStateIsNeverRunnable() {
        assertFalse(PipelineStageGate.canRunStage20(null));
        assertFalse(PipelineStageGate.canRunStage30(null));
    }
}
