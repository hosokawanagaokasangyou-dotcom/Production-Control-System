package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Map;

import org.junit.jupiter.api.Test;

class ReconciliationAppRecordFilterTest {

    private static OrderRecord record(String reqNo, String status, String user) {
        return new OrderRecord(reqNo, status, user, "", "", Map.of(), Map.of());
    }

    @Test
    void recordMatchesFilter_newOnly_matchesStatusContainingShinki() {
        OrderRecord neu =
                record("Y6-24", "新規自動追加 (未登録)", "自動転記");
        OrderRecord existing = record("Y5-5", "既存登録 (相違あり)", "自動転記");
        assertTrue(ReconciliationApp.recordMatchesFilter(neu, "", true));
        assertFalse(ReconciliationApp.recordMatchesFilter(existing, "", true));
    }

    @Test
    void recordMatchesFilter_query_matchesReqNoOrUser() {
        OrderRecord r = record("Y6-24", "新規自動追加 (未登録)", "ｵｶﾓﾄ");
        assertTrue(ReconciliationApp.recordMatchesFilter(r, "y6", false));
        assertTrue(ReconciliationApp.recordMatchesFilter(r, "ｵｶ", false));
        assertFalse(ReconciliationApp.recordMatchesFilter(r, "zzz", false));
    }

    @Test
    void recordMatchesFilter_newOnlyAndQuery_combined() {
        OrderRecord neu = record("Y6-1", "新規自動追加 (未登録)", "A");
        OrderRecord neuOther = record("Y6-2", "新規自動追加 (未登録)", "B");
        assertTrue(ReconciliationApp.recordMatchesFilter(neu, "y6-1", true));
        assertFalse(ReconciliationApp.recordMatchesFilter(neuOther, "y6-1", true));
    }
}
