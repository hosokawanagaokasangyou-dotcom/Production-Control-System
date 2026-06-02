package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class OvertimeSimulationEditStateTest {

    @Test
    void snapOvertimeMinutes_quarterHourSteps() {
        assertEquals(0, OvertimeSimulationEditState.snapOvertimeMinutes(0));
        assertEquals(0, OvertimeSimulationEditState.snapOvertimeMinutes(7));
        assertEquals(15, OvertimeSimulationEditState.snapOvertimeMinutes(8));
        assertEquals(15, OvertimeSimulationEditState.snapOvertimeMinutes(20));
        assertEquals(30, OvertimeSimulationEditState.snapOvertimeMinutes(23));
        assertEquals(720, OvertimeSimulationEditState.snapOvertimeMinutes(800));
    }
}
