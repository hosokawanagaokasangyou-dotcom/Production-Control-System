package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.Test;

class RequestFormMasterCandidateLabelHighlighterTest {

    @Test
    void labelTextFill_matchesFormHeadingColor() {
        assertEquals("#EFF6FF", RequestFormMasterCandidateLabelHighlighter.LABEL_TEXT_FILL_HEX);
    }

    @Test
    void highlightTextFill_isYellow() {
        assertEquals("#FFEB3B", RequestFormMasterCandidateLabelHighlighter.HIGHLIGHT_TEXT_FILL_HEX);
    }

    @Test
    void highlightMask_marksFilterSubstrings() {
        String label = "A2F20AXD0250FN1 | 15020 | 6783 | 1300×250 | EC,梱包";
        boolean[] mask =
                RequestFormMasterCandidateLabelHighlighter.highlightMask(
                        label, List.of("15020", "6783", "EC"));

        assertTrue(mask[label.indexOf("15020")]);
        assertTrue(mask[label.indexOf("6783")]);
        assertTrue(mask[label.indexOf("EC")]);
        assertFalse(mask[0]);
    }

    @Test
    void highlightMask_lowercaseFilter_matchesUppercaseLabel() {
        String label = "A2F20AXD0250FN1 | 15020-NP17 | 6783 | NP17 | 1300×250 | 白 | EC,梱包";
        boolean[] mask =
                RequestFormMasterCandidateLabelHighlighter.highlightMask(
                        label, List.of("np17", "6783", "ec"));

        assertTrue(mask[label.indexOf("NP17")]);
        assertTrue(mask[label.indexOf("6783")]);
        assertTrue(mask[label.indexOf("EC")]);
    }

    @Test
    void highlightMask_emptyFilters_leavesMaskClear() {
        boolean[] mask =
                RequestFormMasterCandidateLabelHighlighter.highlightMask("ABC | 15020", List.of());
        for (boolean on : mask) {
            assertFalse(on);
        }
    }
}
