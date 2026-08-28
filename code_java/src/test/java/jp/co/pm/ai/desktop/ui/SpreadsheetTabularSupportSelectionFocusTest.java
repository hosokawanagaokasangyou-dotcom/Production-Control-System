package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.Set;

import org.junit.jupiter.api.Test;

class SpreadsheetTabularSupportSelectionFocusTest {

    private static final int FIRST_DATA = 1;

    @Test
    void resolveExpansionFocusViewRow_ignoresStaleFocusNotInSelection() {
        assertEquals(
                5,
                SpreadsheetTabularSupport.resolveExpansionFocusViewRow(
                        FIRST_DATA, Set.of(5), 1));
    }

    @Test
    void resolveExpansionFocusViewRow_usesFocusWhenInSelection() {
        assertEquals(
                5,
                SpreadsheetTabularSupport.resolveExpansionFocusViewRow(
                        FIRST_DATA, Set.of(1, 5), 5));
    }

    @Test
    void resolveExpansionFocusViewRow_fallsBackToMaxSelectedWhenFocusNull() {
        assertEquals(
                7,
                SpreadsheetTabularSupport.resolveExpansionFocusViewRow(
                        FIRST_DATA, Set.of(3, 7), null));
    }

    @Test
    void resolveExpansionFocusViewRow_fallsBackToMaxWhenStaleFocusNotInSelection() {
        assertEquals(
                5,
                SpreadsheetTabularSupport.resolveExpansionFocusViewRow(
                        FIRST_DATA, Set.of(1, 5), 3));
    }

    @Test
    void resolveExpansionFocusModelRow_fallsBackToMaxWhenStaleFocusNotInSelection() {
        assertEquals(
                25,
                SpreadsheetTabularSupport.resolveExpansionFocusModelRow(
                        FIRST_DATA, Set.of(17, 25), 19));
    }
}
