package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

class MailHtmlBuilderTest {

    @Test
    @DisplayName("HTML本文は難波の挨拶と拠点表にする")
    void usesNanbaGreetingAndTable() {
        String html = MailHtmlBuilder.build(null, null);
        assertTrue(html.contains("長岡産業の難波です。"), html);
        assertTrue(html.contains("掲題の件、下記にご報告申し上げます。"), html);
        assertTrue(html.contains("B=①+②+③"), html);
    }
}
