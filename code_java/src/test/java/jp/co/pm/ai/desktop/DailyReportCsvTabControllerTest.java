package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Map;

import org.junit.jupiter.api.Test;

class DailyReportCsvTabControllerTest {

    @Test
    void rowMatchesSearch_matchesAnyCell() {
        Map<String, String> row =
                Map.of("依頼NO", "Y8-69", "機械名", "スライス機1", "完了区分", "0:未完");
        assertTrue(DailyReportCsvTabController.rowMatchesSearch(row, "y8-69"));
        assertTrue(DailyReportCsvTabController.rowMatchesSearch(row, "スライス"));
        assertFalse(DailyReportCsvTabController.rowMatchesSearch(row, "UNKNOWN"));
        assertTrue(DailyReportCsvTabController.rowMatchesSearch(row, ""));
        assertTrue(DailyReportCsvTabController.rowMatchesSearch(row, null));
    }
}
