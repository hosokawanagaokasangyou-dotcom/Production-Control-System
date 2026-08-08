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
        List<String> headers =
                List.of(
                        "依頼NO",
                        PlanInputAiSpecialParseColumn.SOURCE_COLUMN_TITLE,
                        PlanInputAiSpecialParseColumn.COLUMN_TITLE);
        List<String> row = new ArrayList<>(List.of("Y3-26", "配台は9/1以降に", "{\"start_date\":\"2026-08-30\"}"));

        assertTrue(
                PlanInputAiSpecialParseColumn.clearStaleParseAfterRemarkEdit(
                        headers, row, PlanInputAiSpecialParseColumn.SOURCE_COLUMN_TITLE));
        assertEquals("", row.get(2));
        assertEquals("配台は9/1以降に", row.get(1));
    }

    @Test
    void renamedRemarkColumnAlsoClearsParseCell() {
        List<String> headers = List.of("依頼NO", "納期回答_備考", "AI納期回答_解析");
        List<String> row = new ArrayList<>(List.of("Y3-26", "8/30以降", "{\"priority\":1}"));

        assertTrue(
                PlanInputAiSpecialParseColumn.clearStaleParseAfterRemarkEdit(
                        headers, row, "納期回答_備考"));
        assertEquals("", row.get(2));
    }

    @Test
    void isParseColumnAcceptsAliases() {
        assertTrue(PlanInputAiSpecialParseColumn.isParseColumn("AI納期回答_解析"));
        assertTrue(PlanInputAiSpecialParseColumn.isParseColumn("AI特別指定_解析"));
        assertFalse(PlanInputAiSpecialParseColumn.isParseColumn("特別指定_備考"));
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
