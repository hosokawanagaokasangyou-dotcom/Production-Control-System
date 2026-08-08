package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class JapanDateTimeDisplayTest {

    @Test
    void formatSavedAtForDisplay_convertsUtcOffsetToJst() {
        assertEquals(
                "2026/08/08 11:28:56",
                JapanDateTimeDisplay.formatSavedAtForDisplay("2026-08-08T02:28:56.680184+00:00"));
    }

    @Test
    void formatSavedAtForDisplay_treatsLegacyLocalAsUtc() {
        assertEquals(
                "2026/08/08 11:50:18",
                JapanDateTimeDisplay.formatSavedAtForDisplay("2026-08-08T02:50:18.359035"));
    }
}
