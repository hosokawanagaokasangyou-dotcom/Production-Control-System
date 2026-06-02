package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.util.Map;
import java.util.function.Predicate;

import org.junit.jupiter.api.Test;

class ReconciliationAppRecordFilterTest {

    private static final Predicate<OrderRecord> NO_ORIGINAL = r -> false;
    private static final Predicate<OrderRecord> HAS_ORIGINAL = r -> true;

    private static OrderRecord record(String reqNo, String status, String user) {
        return new OrderRecord(reqNo, status, user, "", "", Map.of(), Map.of());
    }

    private static OrderRecord recordWithDb(
            String reqNo, String status, String user, Map<String, String> db) {
        return new OrderRecord(reqNo, status, user, "", "", Map.of(), db);
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

    @Test
    void isJuchuRowWithoutRequestFormOriginal_matchesJuchuOnlyRows() {
        OrderRecord juchuOnly =
                recordWithDb(
                        "Y5-1",
                        "既存登録 (原本未確認)",
                        "A",
                        Map.of("入力日", "2026-05-28"));
        OrderRecord withOriginal =
                recordWithDb("Y5-2", "既存登録 (相違あり)", "A", Map.of("入力日", "2026-05-27"));
        OrderRecord shinki =
                recordWithDb(
                        "Y6-1",
                        "新規自動追加 (未登録)",
                        "A",
                        Map.of("入力日", "2026-06-01"));
        assertTrue(
                ReconciliationApp.isJuchuRowWithoutRequestFormOriginal(juchuOnly, NO_ORIGINAL));
        assertFalse(
                ReconciliationApp.isJuchuRowWithoutRequestFormOriginal(withOriginal, HAS_ORIGINAL));
        assertFalse(
                ReconciliationApp.isJuchuRowWithoutRequestFormOriginal(shinki, NO_ORIGINAL));
    }

    @Test
    void recordIncludedInListFilter_modes() {
        OrderRecord existing = record("Y5-5", "既存登録 (相違あり)", "A");
        OrderRecord neu = record("Y6-1", "新規自動追加 (未登録)", "A");
        OrderRecord juchuOnly =
                recordWithDb("Y5-9", "既存登録 (原本未確認)", "A", Map.of("入力日", "2026-05-01"));
        assertTrue(
                ReconciliationApp.recordIncludedInListFilter(
                        existing, ReconciliationApp.RecordListFilterMode.ALL, HAS_ORIGINAL));
        assertTrue(
                ReconciliationApp.recordIncludedInListFilter(
                        neu, ReconciliationApp.RecordListFilterMode.ALL, NO_ORIGINAL));
        assertTrue(
                ReconciliationApp.recordIncludedInListFilter(
                        existing,
                        ReconciliationApp.RecordListFilterMode.EXISTING_ONLY,
                        NO_ORIGINAL));
        assertFalse(
                ReconciliationApp.recordIncludedInListFilter(
                        neu, ReconciliationApp.RecordListFilterMode.EXISTING_ONLY, NO_ORIGINAL));
        assertTrue(
                ReconciliationApp.recordIncludedInListFilter(
                        juchuOnly,
                        ReconciliationApp.RecordListFilterMode.EXISTING_ONLY,
                        NO_ORIGINAL));
        assertTrue(
                ReconciliationApp.recordIncludedInListFilter(
                        existing, ReconciliationApp.RecordListFilterMode.NEW_ONLY, HAS_ORIGINAL));
        assertFalse(
                ReconciliationApp.recordIncludedInListFilter(
                        existing, ReconciliationApp.RecordListFilterMode.NEW_ONLY, NO_ORIGINAL));
    }

    @Test
    void compareRecordByInputDateDesc_newestFirst() {
        OrderRecord older =
                recordWithDb("A", "既存登録 (原本未確認)", "", Map.of("入力日", "2026-01-01"));
        OrderRecord newer =
                recordWithDb("B", "既存登録 (原本未確認)", "", Map.of("入力日", "2026/05/28"));
        assertTrue(ReconciliationApp.compareRecordByInputDateDesc(older, newer) > 0);
        assertEquals(
                LocalDate.of(2026, 5, 28),
                ReconciliationApp.parseInputDateForSort("2026/05/28"));
    }
}
