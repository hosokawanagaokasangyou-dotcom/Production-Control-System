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
    void limitToDateWindow_keepsOnlyInclusiveRange() throws Exception {
        String json =
                """
                {"format_version":1,"ok":true,"members":["A"],
                "dates":["2026-05-20","2026-05-25","2026-06-30"],
                "cells":{"2026-05-20":{"A":{"is_working":true,"eligible_for_assignment":true,
                "overtime_minutes":0,"weekend":false}},
                "2026-05-25":{"A":{"is_working":true,"eligible_for_assignment":true,
                "overtime_minutes":10,"weekend":false}},
                "2026-06-30":{"A":{"is_working":false,"eligible_for_assignment":false,
                "overtime_minutes":0,"weekend":false}}}}
                """;
        AttendanceOvertimePreview.Preview p = AttendanceOvertimePreview.parseJson(json);
        AttendanceOvertimePreview.Preview limited =
                AttendanceOvertimePreview.limitToDateWindow(
                        p, LocalDate.of(2026, 5, 24), LocalDate.of(2026, 5, 26));
        assertEquals(List.of(LocalDate.of(2026, 5, 25)), limited.dates());
        assertEquals(10, limited.cells().get(LocalDate.of(2026, 5, 25)).get("A").overtimeMinutes());
        assertFalse(limited.cells().containsKey(LocalDate.of(2026, 5, 20)));
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
