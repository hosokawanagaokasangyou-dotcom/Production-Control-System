package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class MasterReadSummaryTabControllerTest {

    @Test
    void extractJsonPayload_prefersLastJsonLineWhenLogsPrepended() {
        String in =
                "2026-01-01 12:00:00,123 - INFO - hello\n"
                        + "{\"ok\":true,\"resolved_path\":\"/tmp/m.xlsm\"}\n";
        assertEquals(
                "{\"ok\":true,\"resolved_path\":\"/tmp/m.xlsm\"}",
                MasterReadSummaryTabController.extractJsonPayload(in));
    }

    @Test
    void extractJsonPayload_plainJsonUnchanged() {
        String j = "{\"ok\":false}";
        assertEquals(j, MasterReadSummaryTabController.extractJsonPayload(j));
    }
}
