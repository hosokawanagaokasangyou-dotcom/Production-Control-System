package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

class MailHtmlBuilderTest {

    @Test
    @DisplayName("HTML本文に難波の自己紹介を入れない")
    void omitsNanbaIntro() {
        String html = MailHtmlBuilder.build(null, null);
        assertFalse(html.contains("難波"), html);
        assertFalse(html.contains("長岡産業の難波です。"), html);
        assertTrue(html.contains("いつも大変お世話になっております。"), html);
        assertTrue(html.contains("掲題の件、下記にご報告申し上げます。"), html);
    }
}
