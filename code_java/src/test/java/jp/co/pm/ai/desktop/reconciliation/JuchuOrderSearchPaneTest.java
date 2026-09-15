package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class JuchuOrderSearchPaneTest {

    @Test
    void shouldDisableOpenInspectionSheet_whenNothingSelected() {
        assertTrue(JuchuOrderSearchPane.shouldDisableOpenInspectionSheet(-1, false));
    }

    @Test
    void shouldDisableOpenInspectionSheet_falseWhenRowSelected() {
        assertFalse(JuchuOrderSearchPane.shouldDisableOpenInspectionSheet(0, false));
        assertFalse(JuchuOrderSearchPane.shouldDisableOpenInspectionSheet(3, false));
    }

    @Test
    void shouldDisableOpenInspectionSheet_whenOpenBusyEvenIfSelected() {
        assertTrue(JuchuOrderSearchPane.shouldDisableOpenInspectionSheet(0, true));
    }
}
