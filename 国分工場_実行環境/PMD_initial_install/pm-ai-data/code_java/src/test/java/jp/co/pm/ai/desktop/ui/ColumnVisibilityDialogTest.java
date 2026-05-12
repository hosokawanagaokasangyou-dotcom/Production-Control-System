package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
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

    @Test
    void normalizeMandatoryMask_nullOrNonPositiveCountReturnsNull() {
        assertNull(ColumnVisibilityDialog.normalizeMandatoryMask(null, 3));
        assertNull(ColumnVisibilityDialog.normalizeMandatoryMask(new boolean[] {true}, 0));
    }

    @Test
    void normalizeMandatoryMask_padsShortAndTruncatesLong() {
        boolean[] padded = ColumnVisibilityDialog.normalizeMandatoryMask(new boolean[] {true, false}, 4);
        assertNotNull(padded);
        assertTrue(padded.length == 4 && padded[0] && !padded[1] && !padded[2] && !padded[3]);
        boolean[] cut = ColumnVisibilityDialog.normalizeMandatoryMask(new boolean[] {true, false, true}, 2);
        assertNotNull(cut);
        assertTrue(cut.length == 2 && cut[0] && !cut[1]);
    }

    @Test
    void mergeMandatoryIntoVisibility_visLongerThanMask_fillsFromMaskIndicesOnly() {
        boolean[] vis = {false, false, false, true};
        boolean[] mask = {true, false};
        boolean[] out = ColumnVisibilitySupport.mergeMandatoryIntoVisibility(vis, mask);
        assertNotNull(out);
        assertTrue(out[0] && !out[1] && !out[2] && out[3]);
    }
}
