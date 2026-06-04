package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class PostProcessingProductMasterKubunChoicesTest {

    @Test
    void hinKubun_resolvesCodes() {
        assertEquals("製品", PostProcessingProductMasterKubunChoices.resolveLabel("品区分", "1"));
        assertEquals(
                "1",
                PostProcessingProductMasterKubunChoices.resolveCodeFromPickerInput(
                        "品区分", "1:製品"));
    }

    @Test
    void tenkaiKubun_hasThreeOptions() {
        assertTrue(PostProcessingProductMasterKubunChoices.hasChoices("展開区分"));
        assertEquals(3, PostProcessingProductMasterKubunChoices.pickerLabels("展開区分").size());
    }

    @Test
    void kakoTankaKubun_resolvesZeroOne() {
        assertTrue(PostProcessingProductMasterKubunChoices.hasChoices("加工単価区分"));
        assertEquals(
                "積上",
                PostProcessingProductMasterKubunChoices.resolveLabel("加工単価区分", "0"));
        assertEquals(
                "1",
                PostProcessingProductMasterKubunChoices.resolveCodeFromPickerInput(
                        "加工単価区分", "1:打換"));
    }

    @Test
    void jishaKakoKbn_resolvesZeroOne() {
        assertTrue(PostProcessingProductMasterKubunChoices.hasChoices("自社後加工区分"));
        assertEquals(
                "後加工",
                PostProcessingProductMasterKubunChoices.resolveLabel("自社後加工区分", "1"));
        assertEquals(
                "0",
                PostProcessingProductMasterKubunChoices.resolveCodeFromPickerInput(
                        "自社後加工区分", "0:自社加工"));
        assertEquals(
                "1",
                PostProcessingProductMasterKubunChoices.resolveCodeFromPickerInput(
                        "自社後加工区分", "1:後加工"));
    }

    @Test
    void trimming_resolvesZeroOne() {
        assertEquals("なし", PostProcessingProductMasterKubunChoices.resolveLabel("トリミング", "0"));
        assertEquals(
                "1",
                PostProcessingProductMasterKubunChoices.resolveCodeFromPickerInput(
                        "トリミング", "1:あり"));
    }

    @Test
    void ecSide_hasBundledLabels() {
        assertTrue(PostProcessingProductMasterKubunChoices.hasChoices("EC面指定コード"));
        assertTrue(
                PostProcessingProductMasterKubunChoices.pickerLabels("EC面指定コード").stream()
                        .anyMatch(s -> s.contains("Ｈ面")));
    }
}
