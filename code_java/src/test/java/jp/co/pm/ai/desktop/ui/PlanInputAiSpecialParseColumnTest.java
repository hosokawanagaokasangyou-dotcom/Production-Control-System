package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.ArrayList;
import java.util.List;

import org.junit.jupiter.api.Test;

class PlanInputAiSpecialParseColumnTest {

    private static final List<String> HEADERS =
            List.of(
                    "依頼NO",
                    PlanInputAiSpecialParseColumn.SOURCE_COLUMN_TITLE,
                    PlanInputAiSpecialParseColumn.COLUMN_TITLE);

    @Test
    void remarkEditClearsStaleParseOnTheSameRow() {
        List<String> row = new ArrayList<>(List.of("Y3-26", "配台は9/1以降に", "{\"start_date\":\"2026-08-30\"}"));

        assertTrue(
                PlanInputAiSpecialParseColumn.clearStaleParseAfterRemarkEdit(
                        HEADERS, row, PlanInputAiSpecialParseColumn.SOURCE_COLUMN_TITLE));
        assertEquals("", row.get(2));
        assertEquals("配台は9/1以降に", row.get(1));
    }

    @Test
    void editingOtherColumnsKeepsParseValue() {
        List<String> row = new ArrayList<>(List.of("Y3-26", "配台は9/1以降に", "{\"priority\":1}"));

        assertFalse(
                PlanInputAiSpecialParseColumn.clearStaleParseAfterRemarkEdit(
                        HEADERS, row, "依頼NO"));
        assertEquals("{\"priority\":1}", row.get(2));
    }

    @Test
    void alreadyEmptyParseCellIsNotReportedAsCleared() {
        List<String> row = new ArrayList<>(List.of("Y3-26", "配台は9/1以降に", ""));

        assertFalse(
                PlanInputAiSpecialParseColumn.clearStaleParseAfterRemarkEdit(
                        HEADERS, row, PlanInputAiSpecialParseColumn.SOURCE_COLUMN_TITLE));
    }

    @Test
    void missingParseColumnOrShortRowIsHarmless() {
        List<String> row = new ArrayList<>(List.of("Y3-26", "配台は9/1以降に"));

        assertFalse(
                PlanInputAiSpecialParseColumn.clearStaleParseAfterRemarkEdit(
                        HEADERS, row, PlanInputAiSpecialParseColumn.SOURCE_COLUMN_TITLE));
        assertFalse(
                PlanInputAiSpecialParseColumn.clearStaleParseAfterRemarkEdit(
                        List.of("依頼NO"),
                        new ArrayList<>(List.of("Y3-26")),
                        PlanInputAiSpecialParseColumn.SOURCE_COLUMN_TITLE));
        assertFalse(
                PlanInputAiSpecialParseColumn.clearStaleParseAfterRemarkEdit(null, row, null));
    }
}
