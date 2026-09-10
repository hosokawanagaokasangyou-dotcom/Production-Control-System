package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;

class SpreadsheetPlanInputCellEditSupportChromeSuppressTest {

    @AfterEach
    void tearDown() {
        while (SpreadsheetPlanInputCellEditSupport.isCellEditDialogOpen()) {
            SpreadsheetPlanInputCellEditSupport.endCellEditDialog();
        }
        SpreadsheetPlanInputCellEditSupport.clearChromeRelayoutGraceForTests();
    }

    @Test
    void shouldSkipChromeRelayout_whileDialogOpen() {
        assertFalse(SpreadsheetPlanInputCellEditSupport.shouldSkipChromeRelayout());
        SpreadsheetPlanInputCellEditSupport.beginCellEditDialog();
        assertTrue(SpreadsheetPlanInputCellEditSupport.shouldSkipChromeRelayout());
        SpreadsheetPlanInputCellEditSupport.endCellEditDialog();
        assertTrue(
                SpreadsheetPlanInputCellEditSupport.shouldSkipChromeRelayout(),
                "閉じた直後の grace でも抑止");
    }
}
