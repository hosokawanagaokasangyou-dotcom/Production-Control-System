package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class RequestFormPipelineStatusServiceInHouseTest {

    @Test
    void isInHouseSelfProcessingIraiNo_trueWhenStartsWithTwo() {
        assertTrue(RequestFormPipelineStatusService.isInHouseSelfProcessingIraiNo("2125-02-16"));
        assertTrue(RequestFormPipelineStatusService.isInHouseSelfProcessingIraiNo(" 2125-03-27 "));
    }

    @Test
    void isInHouseSelfProcessingIraiNo_falseOtherwise() {
        assertFalse(RequestFormPipelineStatusService.isInHouseSelfProcessingIraiNo("C8-11"));
        assertFalse(RequestFormPipelineStatusService.isInHouseSelfProcessingIraiNo(""));
        assertFalse(RequestFormPipelineStatusService.isInHouseSelfProcessingIraiNo(null));
    }
}
