package jp.co.pm.ai.desktop.reconciliation;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ReconciliationAppJuchuFileReselectTest {

    @Test
    void juchuFileReselect_reloadsRelatedDataWithoutMasterProductList() {
        assertTrue(ReconciliationApp.shouldReloadRelatedDataAfterJuchuFileReselect());
        assertFalse(ReconciliationApp.shouldReloadMasterProductListAfterJuchuFileReselect());
        String reason = ReconciliationApp.reloadReasonAfterJuchuFileReselect();
        assertTrue(reason.contains("受注ファイル"));
        assertTrue(reason.contains("再読込"));
    }
}
