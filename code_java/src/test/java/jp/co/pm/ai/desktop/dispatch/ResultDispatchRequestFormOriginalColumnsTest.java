package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class ResultDispatchRequestFormOriginalColumnsTest {

    @Test
    void recognizesOriginalDerivedHeaders() {
        assertTrue(ResultDispatchRequestFormOriginalColumns.isDerivedFromRequestFormOriginal("依頼NO"));
        assertTrue(ResultDispatchRequestFormOriginalColumns.isDerivedFromRequestFormOriginal("原反投入日"));
        assertTrue(ResultDispatchRequestFormOriginalColumns.isDerivedFromRequestFormOriginal("品名(製品)"));
    }

    @Test
    void rejectsDispatchAndAladdinPrimaryHeaders() {
        assertFalse(
                ResultDispatchRequestFormOriginalColumns.isDerivedFromRequestFormOriginal("配台試行順番"));
        assertFalse(ResultDispatchRequestFormOriginalColumns.isDerivedFromRequestFormOriginal("換算数量"));
        assertFalse(ResultDispatchRequestFormOriginalColumns.isDerivedFromRequestFormOriginal("加工開始日時"));
        assertFalse(ResultDispatchRequestFormOriginalColumns.isDerivedFromRequestFormOriginal("配台日"));
    }
}
