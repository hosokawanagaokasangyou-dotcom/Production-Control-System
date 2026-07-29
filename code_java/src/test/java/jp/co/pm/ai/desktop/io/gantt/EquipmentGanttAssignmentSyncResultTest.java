package jp.co.pm.ai.desktop.io.gantt;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class EquipmentGanttAssignmentSyncResultTest {

    @Test
    void parseWarningsResponse() throws Exception {
        String json =
                """
                {"format_version":1,"ok":false,"status":"warnings",\
                "timeline_hash":"abc","confirm_token":"tok123",\
                "warnings":[{"code":"absent","message":"欠勤","person":"山田"}],\
                "errors":[]}
                """;
        EquipmentGanttAssignmentSyncResult r =
                EquipmentGanttAssignmentSyncResult.parseJson(json);
        assertFalse(r.ok());
        assertTrue(r.hasWarnings());
        assertEquals("warnings", r.status());
        assertEquals("abc", r.timelineHash());
        assertEquals("tok123", r.confirmToken());
        assertEquals(1, r.warnings().size());
    }
}
