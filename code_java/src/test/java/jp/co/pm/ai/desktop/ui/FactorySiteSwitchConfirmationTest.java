package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import jp.co.pm.ai.desktop.config.FactorySite;

import org.junit.jupiter.api.Test;

class FactorySiteSwitchConfirmationTest {

    @Test
    void shouldPromptUser_onlyForInteractiveSwitch() {
        assertTrue(FactorySiteSwitchConfirmation.shouldPromptUser(false));
        assertFalse(
                FactorySiteSwitchConfirmation.shouldPromptUser(true),
                "起動時の工場復元は確認ダイアログを出さない");
    }

    @Test
    void contentText_includesFromAndToFactoryLabels() {
        String text =
                FactorySiteSwitchConfirmation.contentText(FactorySite.KONAN, FactorySite.KOKUBU);
        assertTrue(text.contains(FactorySite.KONAN.displayLabelJa()));
        assertTrue(text.contains(FactorySite.KOKUBU.displayLabelJa()));
        assertTrue(text.contains("よろしいですか"));
    }

    @Test
    void contentText_whenFromIsNull_usesUnsetLabel() {
        String text = FactorySiteSwitchConfirmation.contentText(null, FactorySite.KONAN);
        assertTrue(text.contains("未設定") || text.contains(FactorySite.KONAN.displayLabelJa()));
        assertEquals(FactorySiteSwitchConfirmation.TITLE, "工場切替の確認");
    }
}
