package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.List;

import org.junit.jupiter.api.Test;

class SpreadsheetTabularSupportChromeRelayoutTest {

    @Test
    void extraFixedRowsToRestore_keepsTitleRowsAndDropsFilterRow() {
        assertEquals(
                List.of(1, 2),
                SpreadsheetTabularSupport.extraFixedRowsToRestore(
                        List.of(
                                SpreadsheetTabularSupport.SPREADSHEET_FILTER_ROW,
                                1,
                                2)));
    }

    @Test
    void extraFixedRowsToRestore_emptyWhenOnlyFilterRow() {
        assertEquals(
                List.of(),
                SpreadsheetTabularSupport.extraFixedRowsToRestore(
                        List.of(SpreadsheetTabularSupport.SPREADSHEET_FILTER_ROW)));
    }
}
