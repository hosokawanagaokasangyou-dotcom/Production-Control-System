package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;

import javafx.application.Platform;
import javafx.scene.control.Label;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

class GlobalAppStatusBarTest {

    @BeforeAll
    static void initJavaFx() {
        try {
            Platform.startup(() -> {});
        } catch (IllegalStateException ignored) {
            // already started
        }
    }

    @Test
    void setMessage_shortensLongLines() {
        Label message = new Label();
        GlobalAppStatusBar bar =
                new GlobalAppStatusBar(message, null, null, null, null, null);
        String longLine = "x".repeat(300);
        bar.setMessage(longLine);
        assertEquals(240, message.getText().length());
        assertEquals('…', message.getText().charAt(239));
    }
}
