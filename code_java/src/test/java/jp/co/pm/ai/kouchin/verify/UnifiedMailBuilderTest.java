package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

class UnifiedMailBuilderTest {

    @Test
    @DisplayName("メール本文に難波の自己紹介を入れない")
    void omitsNanbaIntro() {
        String text = UnifiedMailBuilder.buildText(null, null);
        assertFalse(text.contains("難波"), text);
        assertFalse(text.contains("長岡産業の難波です。"), text);
        assertTrue(text.contains("いつも大変お世話になっております。"), text);
        List<String> lines = UnifiedMailBuilder.buildLines(null, null);
        int greeting = lines.indexOf("いつも大変お世話になっております。");
        int subject = lines.indexOf("掲題の件、下記にご報告申し上げます。");
        assertTrue(greeting >= 0);
        assertTrue(subject > greeting);
        for (int i = greeting + 1; i < subject; i++) {
            assertFalse(lines.get(i).contains("難波"), lines.get(i));
        }
    }
}
