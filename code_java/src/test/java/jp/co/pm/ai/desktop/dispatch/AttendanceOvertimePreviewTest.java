package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

class AttendanceOvertimePreviewTest {

    @Test
    void parseJson_readsCompactPreview() throws Exception {
        String json =
                """
                {"format_version":1,"ok":true,"members":["A"],"dates":["2026-05-23"],
                "cells":{"2026-05-23":{"A":{"is_working":false,"eligible_for_assignment":false,
                "overtime_minutes":0,"weekend":true}}}}
                """;
        AttendanceOvertimePreview.Preview p = AttendanceOvertimePreview.parseJson(json);
        assertTrue(p.ok());
        assertEquals(List.of("A"), p.members());
        assertEquals(LocalDate.of(2026, 5, 23), p.dates().get(0));
        AttendanceOvertimePreview.CellInfo c =
                p.cells().get(LocalDate.of(2026, 5, 23)).get("A");
        assertFalse(c.working());
        assertTrue(c.weekend());
    }

    @Test
    void editState_toggleWorkingAndOvertime() throws Exception {
        String json =
                """
                {"format_version":1,"ok":true,"members":["A"],"dates":["2026-05-23"],
                "cells":{"2026-05-23":{"A":{"is_working":false,"eligible_for_assignment":false,
                "overtime_minutes":0,"weekend":true}}}}
                """;
        AttendanceOvertimePreview.Preview p = AttendanceOvertimePreview.parseJson(json);
        OvertimeSimulationEditState state = new OvertimeSimulationEditState(p);
        LocalDate d = LocalDate.of(2026, 5, 23);
        assertFalse(state.cell(d, "A").currentWorking());
        state.toggleWorking(d, "A");
        assertTrue(state.cell(d, "A").currentWorking());
        state.setOvertimeMinutes(d, "A", 60);
        assertTrue(state.hasChanges());
        var payload = OvertimeSimulationOverridesWriter.buildFromEditState(state);
        assertEquals(Map.of("A", true), payload.workingOverrides().get(d));
        assertEquals(Map.of("A", 60), payload.overtimeMinutes().get(d));
    }
}
