package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.List;
import java.util.Optional;

import org.junit.jupiter.api.Test;

class PlanInputCellEditRoutingTest {

    @Test
    void limitedOperatorColumnUsesChecklistAndOrdinaryColumnsRemainText() {
        assertEquals(
                PlanInputCellEditRouting.Editor.LIMITED_OPERATOR_CHECKLIST,
                PlanInputCellEditRouting.editorFor("担当OP_限定"));
        assertEquals(
                PlanInputCellEditRouting.Editor.TEXT,
                PlanInputCellEditRouting.editorFor("特別指定_備考"));
    }

    @Test
    void aiSpecialParseColumnIsReadOnly() {
        assertEquals(
                PlanInputCellEditRouting.Editor.READ_ONLY,
                PlanInputCellEditRouting.editorFor(
                        PlanInputCellEditRouting.COL_AI_SPECIAL_PARSE));
    }

    @Test
    void cancelledSelectionLeavesCurrentCellValueUnchanged() {
        String current = "[\"山田\"]";

        assertEquals(
                current,
                PlanInputCellEditRouting.applyLimitedOperatorResult(
                        current, Optional.empty()));
        assertEquals(
                "[\"佐藤\",\"山田\"]",
                PlanInputCellEditRouting.applyLimitedOperatorResult(
                        current, Optional.of(List.of("佐藤", "山田"))));
    }
}
