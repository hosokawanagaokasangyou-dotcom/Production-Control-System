package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import javafx.application.Platform;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressBar;

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
                new GlobalAppStatusBar(message, null, null, null, null, null, null, null);
        String longLine = "x".repeat(300);
        bar.setMessage(longLine);
        assertEquals(240, message.getText().length());
        assertEquals('…', message.getText().charAt(239));
    }

    @Test
    void setTaskProgress_repeatedUpdatesKeepBarManagedOnceVisible() {
        ProgressBar barNode = new ProgressBar();
        GlobalAppStatusBar bar =
                new GlobalAppStatusBar(null, null, barNode, null, null, null, null, null);
        bar.setTaskProgress(0.2);
        assertTrue(barNode.isVisible());
        assertTrue(barNode.isManaged());
        bar.setTaskProgress(0.8);
        assertTrue(barNode.isVisible());
        assertTrue(barNode.isManaged());
        assertEquals(0.8, barNode.getProgress(), 1e-9);
        bar.setTaskProgress(null);
        assertFalse(barNode.isVisible());
        assertFalse(barNode.isManaged());
    }
}
