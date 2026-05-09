package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class ColumnVisibilityDialogTest {

    @Test
    void columnTitleMatchesSearch_blankShowsAll() {
        assertTrue(ColumnVisibilityDialog.columnTitleMatchesSearch("colA", ""));
        assertTrue(ColumnVisibilityDialog.columnTitleMatchesSearch("colA", "   "));
        assertTrue(ColumnVisibilityDialog.columnTitleMatchesSearch("colA", null));
    }

    @Test
    void columnTitleMatchesSearch_substring() {
        assertTrue(ColumnVisibilityDialog.columnTitleMatchesSearch("prefixMiddleSuffix", "Middle"));
        assertFalse(ColumnVisibilityDialog.columnTitleMatchesSearch("prefixMiddleSuffix", "Other"));
    }

    @Test
    void columnTitleMatchesSearch_asciiCaseFold() {
        assertTrue(ColumnVisibilityDialog.columnTitleMatchesSearch("RequestNo", "req"));
        assertTrue(ColumnVisibilityDialog.columnTitleMatchesSearch("request", "REQUEST"));
    }
}
