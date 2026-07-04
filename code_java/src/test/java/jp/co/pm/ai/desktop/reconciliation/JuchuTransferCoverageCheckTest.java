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
    void contractNoJuchuStatus_presentWhenMatched() {
        Map<String, String> orig = Map.of("契約Ｎｏ", "ABC-123");
        Map<String, String> juchu = Map.of("契約Ｎｏ", "ABC-123");
        assertEquals("あり", JuchuTransferCoverageCheck.contractNoJuchuStatus(orig, juchu, true));
    }

    @Test
    void contractNoJuchuStatus_missingInJuchu() {
        Map<String, String> orig = Map.of("契約Ｎｏ", "ABC-123");
        assertEquals("なし", JuchuTransferCoverageCheck.contractNoJuchuStatus(orig, Map.of(), true));
    }

    @Test
    void contractNoJuchuStatus_noOriginalContract() {
        assertEquals("-", JuchuTransferCoverageCheck.contractNoJuchuStatus(Map.of(), Map.of(), true));
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
