package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.LinkedHashMap;
import java.util.Map;

import org.junit.jupiter.api.Test;

class JuchuTransferCoverageCheckTest {

    @Test
    void compare_allMatchingFields_fullRate() {
        Map<String, String> orig = new LinkedHashMap<>();
        orig.put("品名", "テスト品");
        orig.put("製品", "ABC-123");
        orig.put("数量1", "1000m");
        orig.put("ユーザー", "テストユーザー");

        Map<String, String> juchu = new LinkedHashMap<>(orig);

        JuchuTransferCoverageCheck.CoverageResult result =
                JuchuTransferCoverageCheck.compare(orig, juchu, null, null);

        assertTrue(result.juchuRowExists());
        assertEquals(4, result.totalWithOriginalValue());
        assertEquals(4, result.matchedCount());
        assertEquals(100.0, result.ratePercent(), 0.01);
        assertEquals("100% (4/4)", result.rateDisplay());
    }

    @Test
    void compare_skipsBlankOriginalValues() {
        Map<String, String> orig = new LinkedHashMap<>();
        orig.put("品名", "テスト品");
        orig.put("製品", "");
        orig.put("数量1", "100");

        Map<String, String> juchu = new LinkedHashMap<>();
        juchu.put("品名", "テスト品");
        juchu.put("製品", "DIFFERENT");
        juchu.put("数量1", "100");

        JuchuTransferCoverageCheck.CoverageResult result =
                JuchuTransferCoverageCheck.compare(orig, juchu, null, null);

        assertEquals(2, result.totalWithOriginalValue());
        assertEquals(2, result.matchedCount());
    }

    @Test
    void compare_noJuchuRow_zeroRate() {
        Map<String, String> orig = new LinkedHashMap<>();
        orig.put("品名", "テスト品");
        orig.put("数量1", "100");

        JuchuTransferCoverageCheck.CoverageResult result =
                JuchuTransferCoverageCheck.compare(orig, Map.of(), null, null);

        assertFalse(result.juchuRowExists());
        assertEquals(0, result.matchedCount());
        assertEquals(0.0, result.ratePercent(), 0.01);
        assertEquals(2, result.mismatchCount());
    }

    @Test
    void compare_dateNormalization() {
        Map<String, String> orig = Map.of("希望納期", "2026/06/15");
        Map<String, String> juchu = Map.of("希望納期", "2026-06-15 00:00:00");

        JuchuTransferCoverageCheck.CoverageResult result =
                JuchuTransferCoverageCheck.compare(orig, juchu, null, null);

        assertEquals(1, result.matchedCount());
    }

    @Test
    void compare_userPartialMatch() {
        Map<String, String> orig = Map.of("ユーザー", "テスト株式会社");
        Map<String, String> juchu = Map.of("ユーザー", "テスト");

        JuchuTransferCoverageCheck.CoverageResult result =
                JuchuTransferCoverageCheck.compare(orig, juchu, null, null);

        assertEquals(1, result.matchedCount());
    }

    @Test
    void formatOriginalContractNoDisplay_multilineJoinedWithSlash() {
        Map<String, String> orig = Map.of("契約Ｎｏ", "186046F\n187062R");
        assertEquals(
                "186046F/187062R",
                JuchuTransferCoverageCheck.formatOriginalContractNoDisplay(orig, true));
    }

    @Test
    void formatOriginalContractNoDisplay_noOriginalShowsDash() {
        assertEquals(
                "-",
                JuchuTransferCoverageCheck.formatOriginalContractNoDisplay(Map.of(), false));
    }

    @Test
    void formatOriginalContractNoDisplay_blankShowsMishuuryoku() {
        assertEquals(
                "未入力",
                JuchuTransferCoverageCheck.formatOriginalContractNoDisplay(Map.of(), true));
    }

    @Test
    void formatJuchuContractNoDisplay_singleValue() {
        Map<String, String> juchu = Map.of("契約Ｎｏ", "186046F");
        assertEquals(
                "186046F",
                JuchuTransferCoverageCheck.formatJuchuContractNoDisplay(juchu, true));
    }

    @Test
    void formatJuchuContractNoDisplay_multilineJoinedWithSlash() {
        Map<String, String> juchu = Map.of("契約Ｎｏ", "186046F\n187062R");
        assertEquals(
                "186046F/187062R",
                JuchuTransferCoverageCheck.formatJuchuContractNoDisplay(juchu, true));
    }

    @Test
    void formatJuchuContractNoDisplay_blankShowsMishuuryoku() {
        assertEquals(
                "未入力",
                JuchuTransferCoverageCheck.formatJuchuContractNoDisplay(Map.of(), true));
        assertEquals(
                "未入力",
                JuchuTransferCoverageCheck.formatJuchuContractNoDisplay(
                        Map.of("契約Ｎｏ", ""), true));
        assertEquals(
                "未入力",
                JuchuTransferCoverageCheck.formatJuchuContractNoDisplay(Map.of(), false));
    }

    @Test
    void mergeContractNoValues_combinesTwoRows() {
        Map<String, String> first = new LinkedHashMap<>();
        first.put("契約Ｎｏ", "186046F");
        Map<String, String> second = new LinkedHashMap<>();
        second.put("契約Ｎｏ", "187062R");
        JuchuTransferCoverageCheck.mergeContractNoValues(first, second);
        assertEquals(
                "186046F/187062R",
                JuchuTransferCoverageCheck.formatJuchuContractNoDisplay(first, true));
    }

    @Test
    void compare_contractNoSlashVsNewlineMatches() {
        Map<String, String> orig = Map.of("契約Ｎｏ", "187065Y/187066S");
        Map<String, String> juchu = Map.of("契約Ｎｏ", "187065Y\n187066S");

        JuchuTransferCoverageCheck.CoverageResult result =
                JuchuTransferCoverageCheck.compare(orig, juchu, null, null);

        assertEquals(1, result.totalWithOriginalValue());
        assertEquals(1, result.matchedCount());
        assertTrue(result.details().get(0).matched());
    }

    @Test
    void compare_inputDateShortMonthDayMatchesFullDate() {
        Map<String, String> orig = Map.of("投入日", "7/15");
        Map<String, String> juchu = Map.of("投入日", "2026-07-15");

        JuchuTransferCoverageCheck.CoverageResult result =
                JuchuTransferCoverageCheck.compare(orig, juchu, null, null);

        assertEquals(1, result.totalWithOriginalValue());
        assertEquals(1, result.matchedCount());
        assertTrue(result.details().get(0).matched());
    }

    @Test
    void compare_kiboNokiJapaneseMonthDayMatchesIsoDate() {
        Map<String, String> orig = Map.of("希望納期", "7月7日");
        Map<String, String> juchu = Map.of("希望納期", "2026-07-07");

        JuchuTransferCoverageCheck.CoverageResult result =
                JuchuTransferCoverageCheck.compare(orig, juchu, null, null);

        assertEquals(1, result.totalWithOriginalValue());
        assertEquals(1, result.matchedCount());
        assertTrue(result.details().get(0).matched());
    }

    @Test
    void compare_genpanQuantityIgnoresThousandsSeparator() {
        Map<String, String> orig = Map.of("原反数量", "1,050");
        Map<String, String> juchu = Map.of("原反数量", "1050");

        JuchuTransferCoverageCheck.CoverageResult result =
                JuchuTransferCoverageCheck.compare(orig, juchu, null, null);

        assertEquals(1, result.totalWithOriginalValue());
        assertEquals(1, result.matchedCount());
        assertTrue(result.details().get(0).matched());
    }

    @Test
    void compare_excludedColumnSkipped() {
        Map<String, String> orig = Map.of("品名", "A", "製品", "B");
        Map<String, String> juchu = Map.of("品名", "A", "製品", "WRONG");

        JuchuHeaderAliasRegistry registry = new JuchuHeaderAliasRegistry();
        String path = "C:/test/juchu.xlsm";
        registry.setExcludedFromTransfer(path, JuchuSheetColumnLayout.Col.SEIHIN);

        JuchuTransferCoverageCheck.CoverageResult result =
                JuchuTransferCoverageCheck.compare(orig, juchu, registry, path);

        assertEquals(1, result.totalWithOriginalValue());
        assertEquals(1, result.matchedCount());
    }
}
