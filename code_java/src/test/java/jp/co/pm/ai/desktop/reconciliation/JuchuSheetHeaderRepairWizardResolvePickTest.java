package jp.co.pm.ai.desktop.reconciliation;

import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;

class JuchuSheetHeaderRepairWizardResolvePickTest {

    private static JuchuSheetColumnLayout.ExcelHeaderPick pick(
            String letter, int index, String header) {
        return new JuchuSheetColumnLayout.ExcelHeaderPick(letter, index, header);
    }

    @Test
    void resolvePick_matchesDisplayLabelExactly() {
        var bu = pick("BU", 72, "商品(製品)");
        var bv = pick("BV", 73, "商品(原反)");
        var picks = List.of(bu, bv);

        var resolved =
                JuchuSheetHeaderRepairWizard.resolvePick("BV列: 商品(原反)", picks);
        assertNotNull(resolved);
        assertEquals("BV", resolved.columnLetter());
        assertEquals("商品(原反)", resolved.headerText());
    }

    @Test
    void resolvePick_prefersColumnLetterWhenHeaderTextIsDuplicate() {
        var bu = pick("BU", 72, "ラベﾙ色&呼称");
        var bv = pick("BV", 73, "ラベﾙ色&呼称");
        var picks = List.of(bu, bv);

        var resolved =
                JuchuSheetHeaderRepairWizard.resolvePick("BV列: ラベﾙ色&呼称", picks);
        assertNotNull(resolved);
        assertEquals("BV", resolved.columnLetter());
    }

    @Test
    void resolvePick_fallsBackToColumnLetterOnly() {
        var bu = pick("BU", 72, "商品(製品)");
        var bv = pick("BV", 73, "商品(原反)");
        var picks = List.of(bu, bv);

        var resolved = JuchuSheetHeaderRepairWizard.resolvePick("BV列: 存在しない見出し", picks);
        assertNotNull(resolved);
        assertEquals("BV", resolved.columnLetter());
    }

    @Test
    void resolvePick_returnsNullForBlank() {
        var picks = List.of(pick("AQ", 42, "ラベﾙ色&呼称"));
        assertNull(JuchuSheetHeaderRepairWizard.resolvePick("", picks));
        assertNull(JuchuSheetHeaderRepairWizard.resolvePick(null, picks));
    }
}
