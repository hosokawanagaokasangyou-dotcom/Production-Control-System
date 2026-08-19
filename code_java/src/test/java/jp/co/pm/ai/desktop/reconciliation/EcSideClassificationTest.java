package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

import java.util.Map;

import jp.co.pm.ai.desktop.reconciliation.RequestFormOriginalIndexLookup;

class EcSideClassificationTest {

    @Test
    void classify_doubleWhenEcMenBlankAndOriginalRef() {
        assertEquals(
                EcSideClassification.DOUBLE_SIDED,
                EcSideClassification.classify("EC,ゴミ除去", "", true, true));
        assertEquals(
                EcSideClassification.DOUBLE_SIDED,
                EcSideClassification.classify("EC,ゴミ除去", null, true, true));
        assertEquals(
                EcSideClassification.DOUBLE_SIDED,
                EcSideClassification.classify("EC,ゴミ除去", "-", true, true));
    }

    @Test
    void classify_unknownWhenEcMenBlankWithoutOriginalRef() {
        assertEquals(
                EcSideClassification.UNKNOWN,
                EcSideClassification.classify("EC,ゴミ除去", "", true, false));
    }

    @Test
    void classify_singleWhenEcMenPresent() {
        assertEquals(EcSideClassification.SINGLE_SIDED, EcSideClassification.classify("EC,ゴミ除去", "H"));
        assertEquals(EcSideClassification.SINGLE_SIDED, EcSideClassification.classify("EC", "Ｈ面"));
    }

    @Test
    void classify_emptyWithoutEcProcess() {
        assertEquals("", EcSideClassification.classify("スリット,SEC", ""));
        assertEquals("", EcSideClassification.classify("スリット,SEC", "両面"));
    }

    @Test
    void processContentHasEc_detectsEcToken() {
        assertTrue(EcSideClassification.processContentHasEc("EC,ゴミ除去"));
        assertTrue(EcSideClassification.processContentHasEc("EC（片面）"));
        assertTrue(EcSideClassification.processContentHasEc("Ec"));
        assertFalse(EcSideClassification.processContentHasEc("スリット,SEC"));
    }

    @Test
    void classify_w816_juchuEcLowercaseProcessContent() {
        assertEquals(
                EcSideClassification.DOUBLE_SIDED,
                EcSideClassification.classify("Ec", "両面"));
    }

    @Test
    void parentIraiNoLookupKey_stripsTrailingNumericBranch() {
        assertEquals("W7-22", EcSideClassification.parentIraiNoLookupKey("W7-22-1"));
        assertEquals("CS-3", EcSideClassification.parentIraiNoLookupKey("CS-3-1"));
        assertEquals("", EcSideClassification.parentIraiNoLookupKey("W7-22"));
    }

    @Test
    void classify_unknownWhenJuchuMissing() {
        assertEquals(
                EcSideClassification.UNKNOWN,
                EcSideClassification.classify("EC,ゴミ除去", "", false));
        assertEquals(
                EcSideClassification.UNKNOWN,
                EcSideClassification.classify("EC,ゴミ除去", "H", false));
    }

    @Test
    void classify_doubleWhenJuchuEcMenRyomen() {
        assertEquals(EcSideClassification.DOUBLE_SIDED, EcSideClassification.classify("EC", "両面"));
    }

    @Test
    void ecDispatchPassCount_secDoubleSidedIsOnePass() {
        assertEquals(
                2,
                EcSideClassification.ecDispatchPassCount(
                        EcSideClassification.DOUBLE_SIDED, "EC", "KONAN"));
        assertEquals(
                1,
                EcSideClassification.ecDispatchPassCount(
                        EcSideClassification.DOUBLE_SIDED, "SEC", "KONAN"));
        assertEquals(
                1,
                EcSideClassification.ecDispatchPassCount(
                        EcSideClassification.DOUBLE_SIDED, "SEC機　湖南", "KONAN"));
        assertEquals(
                1,
                EcSideClassification.ecDispatchPassCount(
                        EcSideClassification.SINGLE_SIDED, "EC", "KONAN"));
    }

    @Test
    void ecDispatchPassCount_kokubuDoubleSidedIsOnePass() {
        assertEquals(
                1,
                EcSideClassification.ecDispatchPassCount(
                        EcSideClassification.DOUBLE_SIDED, "EC", "KOKUBU"));
        assertEquals(
                1,
                EcSideClassification.ecDispatchPassCount(
                        EcSideClassification.DOUBLE_SIDED, "EC機　国分", "KOKUBU"));
        assertEquals(
                1,
                EcSideClassification.ecDispatchPassCount(
                        EcSideClassification.DOUBLE_SIDED, "SEC", "KOKUBU"));
        assertEquals(
                1,
                EcSideClassification.ecDispatchPassCount(
                        EcSideClassification.SINGLE_SIDED, "EC", "KOKUBU"));
    }

    @Test
    void resolveEcSideClass_fallsBackToParentKey() {
        Map<String, String> map =
                Map.of(
                        RequestFormOriginalIndexLookup.normalizeIraiNoKey("W7-22"),
                        EcSideClassification.SINGLE_SIDED);
        assertEquals(
                EcSideClassification.SINGLE_SIDED,
                EcSideClassification.resolveEcSideClass(map, "W7-22-1"));
    }
}
