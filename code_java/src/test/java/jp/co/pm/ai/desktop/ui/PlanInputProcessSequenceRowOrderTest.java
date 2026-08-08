package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.List;

import org.junit.jupiter.api.Test;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

class PlanInputProcessSequenceRowOrderTest {

    @Test
    void stabilize_restoresProcessSequenceWithinTaskId() {
        List<String> headers =
                List.of(
                        PlanInputProcessSequenceRowOrder.COL_DISPATCH_TRIAL_ORDER,
                        PlanInputProcessSequenceRowOrder.COL_TASK_ID,
                        PlanInputProcessSequenceRowOrder.COL_PROCESS,
                        PlanInputProcessSequenceRowOrder.COL_PROCESS_CONTENT);
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        rows.add(row("3", "V9-9", "融着", "スリット,分割,融着"));
        rows.add(row("1", "V9-9", "スリット", "スリット,分割,融着"));
        rows.add(row("2", "V9-9", "分割", "スリット,分割,融着"));
        rows.add(row("4", "W1-1", "EC", "EC,検査"));

        PlanInputProcessSequenceRowOrder.stabilizeAndRenumberDispatchTrialOrder(headers, rows);

        assertEquals("1", cell(rows, 0, 0));
        assertEquals("V9-9", cell(rows, 0, 1));
        assertEquals("スリット", cell(rows, 0, 2));

        assertEquals("2", cell(rows, 1, 0));
        assertEquals("分割", cell(rows, 1, 2));

        assertEquals("3", cell(rows, 2, 0));
        assertEquals("融着", cell(rows, 2, 2));

        assertEquals("4", cell(rows, 3, 0));
        assertEquals("W1-1", cell(rows, 3, 1));
    }

    @Test
    void stabilize_excludeYesRowDoesNotFollowEligibleBlock() {
        List<String> headers =
                List.of(
                        PlanInputProcessSequenceRowOrder.COL_DISPATCH_TRIAL_ORDER,
                        PlanInputProcessSequenceRowOrder.COL_TASK_ID,
                        PlanInputProcessSequenceRowOrder.COL_PROCESS,
                        PlanInputProcessSequenceRowOrder.COL_PROCESS_CONTENT,
                        PlanInputProcessSequenceRowOrder.COL_EXCLUDE_FROM_ASSIGNMENT);
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        rows.add(row("1", "Y6-2", "スライス", "スライス", ""));
        rows.add(row("2", "E6-1", "スリット", "スリット", ""));
        rows.add(row("3", "T6-1", "分割", "分割,エンボス,巻返し", "yes"));
        rows.add(row("4", "T6-1", "エンボス", "分割,エンボス,巻返し", ""));
        rows.add(row("5", "T6-1", "巻返し", "分割,エンボス,巻返し", ""));
        rows.add(row("20", "T6-1", "分割", "分割,エンボス,巻返し", "yes"));

        PlanInputProcessSequenceRowOrder.stabilizeAndRenumberDispatchTrialOrder(headers, rows);

        assertEquals("分割", cell(rows, 2, 2));
        assertEquals("yes", cell(rows, 2, 4));
        assertEquals("エンボス", cell(rows, 3, 2));
        assertEquals("巻返し", cell(rows, 4, 2));
        assertEquals("分割", cell(rows, 5, 2));
        assertEquals("yes", cell(rows, 5, 4));
    }

    @Test
    void moveRowsForUserReorder_movesEligibleBlockTogether_excludesStay() {
        List<String> headers =
                List.of(
                        PlanInputProcessSequenceRowOrder.COL_DISPATCH_TRIAL_ORDER,
                        PlanInputProcessSequenceRowOrder.COL_TASK_ID,
                        PlanInputProcessSequenceRowOrder.COL_PROCESS,
                        PlanInputProcessSequenceRowOrder.COL_PROCESS_CONTENT,
                        PlanInputProcessSequenceRowOrder.COL_EXCLUDE_FROM_ASSIGNMENT);
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        rows.add(row("1", "Y6-2", "スライス", "スライス", ""));
        rows.add(row("2", "E6-1", "スリット", "スリット", ""));
        rows.add(row("3", "T6-1", "分割", "エンボス,巻返し", "yes"));
        rows.add(row("4", "T6-1", "エンボス", "エンボス,巻返し", ""));
        rows.add(row("5", "T6-1", "巻返し", "エンボス,巻返し", ""));
        rows.add(row("6", "V6-3", "スライス", "スライス", ""));

        PlanInputProcessSequenceRowOrder.moveRowsForUserReorder(headers, rows, 3, 5);

        assertEquals("分割", cell(rows, 2, 2));
        assertEquals("yes", cell(rows, 2, 4));
        assertEquals("V6-3", cell(rows, 3, 1));
        assertEquals("T6-1", cell(rows, 4, 1));
        assertEquals("エンボス", cell(rows, 4, 2));
        assertEquals("T6-1", cell(rows, 5, 1));
        assertEquals("巻返し", cell(rows, 5, 2));

        PlanInputProcessSequenceRowOrder.stabilizeAndRenumberDispatchTrialOrder(headers, rows);

        assertEquals("V6-3", cell(rows, 3, 1));
        assertEquals("T6-1", cell(rows, 4, 1));
        assertEquals("エンボス", cell(rows, 4, 2));
        assertEquals("T6-1", cell(rows, 5, 1));
        assertEquals("巻返し", cell(rows, 5, 2));
    }

    @Test
    void stabilize_ordersEligibleBlocksByDispatchTrialOrder() {
        List<String> headers =
                List.of(
                        PlanInputProcessSequenceRowOrder.COL_DISPATCH_TRIAL_ORDER,
                        PlanInputProcessSequenceRowOrder.COL_TASK_ID,
                        PlanInputProcessSequenceRowOrder.COL_PROCESS,
                        PlanInputProcessSequenceRowOrder.COL_PROCESS_CONTENT);
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        rows.add(row("1", "Y6-2", "スライス", "スライス"));
        rows.add(row("2", "E6-1", "スリット", "スリット"));
        rows.add(row("10", "T6-1", "巻返し", "エンボス,巻返し"));
        rows.add(row("11", "T6-1", "エンボス", "エンボス,巻返し"));
        rows.add(row("5", "V6-3", "スライス", "スライス"));

        PlanInputProcessSequenceRowOrder.stabilizeAndRenumberDispatchTrialOrder(headers, rows);

        assertEquals("Y6-2", cell(rows, 0, 1));
        assertEquals("E6-1", cell(rows, 1, 1));
        assertEquals("V6-3", cell(rows, 2, 1));
        assertEquals("T6-1", cell(rows, 3, 1));
        assertEquals("エンボス", cell(rows, 3, 2));
        assertEquals("T6-1", cell(rows, 4, 1));
        assertEquals("巻返し", cell(rows, 4, 2));
    }

    @Test
    void moveBlockUp_multiProcessTask_swapsWithBlockAbove() {
        List<String> headers =
                List.of(
                        PlanInputProcessSequenceRowOrder.COL_DISPATCH_TRIAL_ORDER,
                        PlanInputProcessSequenceRowOrder.COL_TASK_ID,
                        PlanInputProcessSequenceRowOrder.COL_PROCESS,
                        PlanInputProcessSequenceRowOrder.COL_PROCESS_CONTENT,
                        PlanInputProcessSequenceRowOrder.COL_EXCLUDE_FROM_ASSIGNMENT);
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        rows.add(row("1", "Y6-2", "スライス", "スライス", ""));
        rows.add(row("2", "E6-1", "スリット", "スリット", ""));
        rows.add(row("3", "T6-1", "分割", "エンボス,巻返し", "yes"));
        rows.add(row("4", "T6-1", "エンボス", "エンボス,巻返し", ""));
        rows.add(row("5", "T6-1", "巻返し", "エンボス,巻返し", ""));

        PlanInputSpreadsheetRowReorder.moveBlockUp(headers, rows, 4);

        assertEquals("Y6-2", cell(rows, 0, 1));
        assertEquals("T6-1", cell(rows, 1, 1));
        assertEquals("分割", cell(rows, 1, 2));
        assertEquals("yes", cell(rows, 1, 4));
        assertEquals("T6-1", cell(rows, 2, 1));
        assertEquals("エンボス", cell(rows, 2, 2));
        assertEquals("T6-1", cell(rows, 3, 1));
        assertEquals("巻返し", cell(rows, 3, 2));
        assertEquals("E6-1", cell(rows, 4, 1));
    }

    @Test
    void stabilize_sortsByDecimalKeysThenRenumbersToIntegers() {
        List<String> headers =
                List.of(
                        PlanInputProcessSequenceRowOrder.COL_DISPATCH_TRIAL_ORDER,
                        PlanInputProcessSequenceRowOrder.COL_TASK_ID,
                        PlanInputProcessSequenceRowOrder.COL_PROCESS,
                        PlanInputProcessSequenceRowOrder.COL_PROCESS_CONTENT);
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        rows.add(row("3", "A", "EC", "EC"));
        rows.add(row("1.5", "B", "EC", "EC"));
        rows.add(row("1", "C", "EC", "EC"));

        PlanInputProcessSequenceRowOrder.stabilizeAndRenumberDispatchTrialOrder(headers, rows);

        assertEquals("C", cell(rows, 0, 1));
        assertEquals("B", cell(rows, 1, 1));
        assertEquals("A", cell(rows, 2, 1));
        assertEquals("1", cell(rows, 0, 0));
        assertEquals("2", cell(rows, 1, 0));
        assertEquals("3", cell(rows, 2, 0));
    }

    @Test
    void parsePositiveTrialOrderSortKey_acceptsDecimals() {
        assertEquals(1.5, PlanInputProcessSequenceRowOrder.parsePositiveTrialOrderSortKey("1.5"));
        assertEquals(2.0, PlanInputProcessSequenceRowOrder.parsePositiveTrialOrderSortKey("2"));
    }

    @Test
    void processSequenceRank_matchesNormalizedTokens() {
        List<String> tokens = List.of("スリット", "分割", "融着");
        assertEquals(0, PlanInputProcessSequenceRowOrder.processSequenceRank("スリット", tokens));
        assertEquals(1, PlanInputProcessSequenceRowOrder.processSequenceRank("　分割　", tokens));
        assertEquals(2, PlanInputProcessSequenceRowOrder.processSequenceRank("融着", tokens));
    }

    private static ObservableList<String> row(String dto, String tid, String proc, String content) {
        return FXCollections.observableArrayList(dto, tid, proc, content);
    }

    private static ObservableList<String> row(
            String dto, String tid, String proc, String content, String exclude) {
        return FXCollections.observableArrayList(dto, tid, proc, content, exclude);
    }

    private static String cell(ObservableList<ObservableList<String>> rows, int r, int c) {
        return rows.get(r).get(c);
    }
}
