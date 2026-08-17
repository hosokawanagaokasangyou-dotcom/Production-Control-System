package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class ThemedAlertContentSupportTest {

    @Test
    void needsScrollableContent_falseForShortMessage() {
        assertFalse(ThemedAlertContentSupport.needsScrollableContent("差異 1件"));
        assertFalse(ThemedAlertContentSupport.needsScrollableContent(null));
        assertFalse(ThemedAlertContentSupport.needsScrollableContent(""));
    }

    @Test
    void needsScrollableContent_trueWhenManyLines() {
        String body = "差異 20件\n\n" + "機械=スリット機1 湖南 依頼NO=JR1 工程=スリット\n".repeat(12);
        assertTrue(ThemedAlertContentSupport.needsScrollableContent(body));
    }
}
