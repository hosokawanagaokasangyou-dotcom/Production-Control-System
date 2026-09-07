package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import org.junit.jupiter.api.Test;

class JuchuOrderSearchTest {

    private static OrderRecord rec(String reqNo, Map<String, String> db) {
        return new OrderRecord(reqNo, "既存", "U", "", "", Map.of(), db);
    }

    @Test
    void validation_requiresDateRangeAndAtLeastOneKeyword() {
        assertTrue(
                new JuchuOrderSearchCriteria(null, LocalDate.of(2026, 6, 1), "A", "")
                        .validationError()
                        .isPresent());
        assertTrue(
                new JuchuOrderSearchCriteria(LocalDate.of(2026, 6, 1), null, "A", "")
                        .validationError()
                        .isPresent());
        assertTrue(
                new JuchuOrderSearchCriteria(
                                LocalDate.of(2026, 6, 2), LocalDate.of(2026, 6, 1), "A", "")
                        .validationError()
                        .isPresent());
        assertTrue(
                new JuchuOrderSearchCriteria(
                                LocalDate.of(2026, 6, 1), LocalDate.of(2026, 6, 30), "", "  ")
                        .validationError()
                        .isPresent());
        assertEquals(
                Optional.empty(),
                new JuchuOrderSearchCriteria(
                                LocalDate.of(2026, 6, 1), LocalDate.of(2026, 6, 30), "製品X", "")
                        .validationError());
    }

    @Test
    void matches_hopeDeliveryInRange() {
        var c =
                new JuchuOrderSearchCriteria(
                        LocalDate.of(2026, 6, 1), LocalDate.of(2026, 6, 30), "ABC", "");
        assertTrue(
                JuchuOrderSearch.matches(
                        rec("1", Map.of("希望納期", "2026-06-15", "製品", "XXABCXX")), c));
        assertFalse(
                JuchuOrderSearch.matches(
                        rec("2", Map.of("希望納期", "2026-07-01", "製品", "XXABCXX")), c));
    }

    @Test
    void matches_adjustDeliveryInRange_evenIfHopeOutside() {
        var c =
                new JuchuOrderSearchCriteria(
                        LocalDate.of(2026, 6, 1), LocalDate.of(2026, 6, 30), "ABC", "");
        assertTrue(
                JuchuOrderSearch.matches(
                        rec(
                                "1",
                                Map.of(
                                        "希望納期",
                                        "2026-05-01",
                                        "調整納期",
                                        "2026-06-10",
                                        "製品",
                                        "ABC")),
                        c));
    }

    @Test
    void matches_productOrRawMaterial_or() {
        var c =
                new JuchuOrderSearchCriteria(
                        LocalDate.of(2026, 6, 1), LocalDate.of(2026, 6, 30), "PROD", "RAW");
        assertTrue(
                JuchuOrderSearch.matches(
                        rec("1", Map.of("希望納期", "2026-06-10", "製品", "xxPRODyy", "品名1", "zzz")),
                        c));
        assertTrue(
                JuchuOrderSearch.matches(
                        rec("2", Map.of("希望納期", "2026-06-10", "製品", "zzz", "原反品名", "xxRAWyy")),
                        c));
        assertFalse(
                JuchuOrderSearch.matches(
                        rec("3", Map.of("希望納期", "2026-06-10", "製品", "zzz", "品名1", "zzz")),
                        c));
    }

    @Test
    void matches_unparseableDelivery_doesNotHitOnDate() {
        var c =
                new JuchuOrderSearchCriteria(
                        LocalDate.of(2026, 6, 1), LocalDate.of(2026, 6, 30), "ABC", "");
        assertFalse(
                JuchuOrderSearch.matches(
                        rec("1", Map.of("希望納期", "不明", "調整納期", "", "製品", "ABC")), c));
    }

    @Test
    void filter_rejectsInvalidCriteria() {
        var invalid =
                new JuchuOrderSearchCriteria(
                        LocalDate.of(2026, 6, 1), LocalDate.of(2026, 6, 30), "", "");
        IllegalArgumentException ex =
                assertThrows(
                        IllegalArgumentException.class,
                        () -> JuchuOrderSearch.filter(List.of(), invalid));
        assertEquals("製品名または投入原反を入力してください", ex.getMessage());
    }

    @Test
    void filter_returnsMatchingOnly() {
        var c =
                new JuchuOrderSearchCriteria(
                        LocalDate.of(2026, 6, 1), LocalDate.of(2026, 6, 30), "HIT", "");
        List<OrderRecord> out =
                JuchuOrderSearch.filter(
                        List.of(
                                rec("a", Map.of("希望納期", "2026-06-05", "製品", "HIT")),
                                rec("b", Map.of("希望納期", "2026-06-05", "製品", "MISS"))),
                        c);
        assertEquals(1, out.size());
        assertEquals("a", out.get(0).getReqNo());
    }
}
