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
        assertEquals(
                Optional.empty(),
                new JuchuOrderSearchCriteria(
                                LocalDate.of(2026, 6, 1),
                                LocalDate.of(2026, 6, 30),
                                "",
                                "",
                                "スライス",
                                "")
                        .validationError());
        assertEquals(
                Optional.empty(),
                new JuchuOrderSearchCriteria(
                                LocalDate.of(2026, 6, 1),
                                LocalDate.of(2026, 6, 30),
                                "",
                                "",
                                "",
                                "SEC")
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
        assertEquals("製品名・投入原反・機械名・工程名のいずれかを入力してください", ex.getMessage());
    }

    @Test
    void matches_machineName_partial() {
        var c =
                new JuchuOrderSearchCriteria(
                        LocalDate.of(2026, 6, 1),
                        LocalDate.of(2026, 6, 30),
                        "",
                        "",
                        "スライス",
                        "");
        assertTrue(
                JuchuOrderSearch.matches(
                        rec(
                                "1",
                                Map.of(
                                        "希望納期",
                                        "2026-06-10",
                                        "機械名",
                                        "スライス機1 湖南")),
                        c));
        assertTrue(
                JuchuOrderSearch.matches(
                        rec("2", Map.of("希望納期", "2026-06-10", "機械", "スライス2")), c));
        assertFalse(
                JuchuOrderSearch.matches(
                        rec("3", Map.of("希望納期", "2026-06-10", "製品", "スライス製品")), c));
    }

    @Test
    void matches_processName_orKakouNaiyo() {
        var c =
                new JuchuOrderSearchCriteria(
                        LocalDate.of(2026, 6, 1), LocalDate.of(2026, 6, 30), "", "", "", "SEC");
        assertTrue(
                JuchuOrderSearch.matches(
                        rec("1", Map.of("希望納期", "2026-06-10", "工程名", "SEC工程")), c));
        assertTrue(
                JuchuOrderSearch.matches(
                        rec("2", Map.of("希望納期", "2026-06-10", "加工内容", "①SEC")), c));
        assertFalse(
                JuchuOrderSearch.matches(
                        rec("3", Map.of("希望納期", "2026-06-10", "製品", "SECフィルム")), c));
    }

    @Test
    void matches_machineAndProcess_fromExtraHaystack() {
        var c =
                new JuchuOrderSearchCriteria(
                        LocalDate.of(2026, 6, 1),
                        LocalDate.of(2026, 6, 30),
                        "",
                        "",
                        "W9",
                        "スリット");
        OrderRecord rec = rec("C8-9", Map.of("希望納期", "2026-06-10", "製品", "X"));
        assertFalse(JuchuOrderSearch.matches(rec, c));
        assertTrue(JuchuOrderSearch.matches(rec, c, "W9-1 湖南", "スリット カット"));
        assertFalse(JuchuOrderSearch.matches(rec, c, "W9-1 湖南", "SEC"));
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

    @Test
    void productCandidates_uniqueSplitLinesIgnoreBlank() {
        List<String> names =
                JuchuOrderSearch.productCandidates(
                        List.of(
                                rec("1", Map.of("製品", "B製品")),
                                rec("2", Map.of("製品", "A製品\nC製品")),
                                rec("3", Map.of("製品", "A製品")),
                                rec("4", Map.of("製品", "  "))));
        assertEquals(List.of("A製品", "B製品", "C製品"), names);
    }

    @Test
    void rawMaterialCandidates_prefersHinmei1() {
        List<String> names =
                JuchuOrderSearch.rawMaterialCandidates(
                        List.of(
                                rec("1", Map.of("品名1", "原反A")),
                                rec("2", Map.of("原反品名", "原反B")),
                                rec("3", Map.of("品名1", "原反A"))));
        assertEquals(List.of("原反A", "原反B"), names);
    }

    @Test
    void machineCandidates_mergesDbAndExtras() {
        List<String> names =
                JuchuOrderSearch.machineCandidates(
                        List.of(rec("1", Map.of("機械名", "スライス機1 湖南"))),
                        List.of("W9-1", "スライス機1 湖南", ""));
        assertEquals(2, names.size());
        assertTrue(names.contains("スライス機1 湖南"));
        assertTrue(names.contains("W9-1"));
    }

    @Test
    void processCandidates_usesKouteiAndKakouNaiyo() {
        List<String> names =
                JuchuOrderSearch.processCandidates(
                        List.of(
                                rec("1", Map.of("工程名", "SEC")),
                                rec("2", Map.of("加工内容", "スリット"))),
                        List.of("カット"));
        assertEquals(3, names.size());
        assertTrue(names.contains("SEC"));
        assertTrue(names.contains("スリット"));
        assertTrue(names.contains("カット"));
    }
}
