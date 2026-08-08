package jp.co.pm.ai.desktop.reconciliation;

import org.junit.jupiter.api.Test;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ReconciliationAppProvisionalReloadTest {

    @Test
    void buildProvisionalRecordsFromJuchu_marksRowsAsReconciling() {
        Map<String, String> dbRow = new LinkedHashMap<>();
        dbRow.put("依頼No", "JR001");
        dbRow.put("ユーザー", "テスト");
        dbRow.put("製品", "P1");
        Map<String, Map<String, String>> dbRows = Map.of("jr001", dbRow);

        List<OrderRecord> records = ReconciliationApp.buildProvisionalRecordsFromJuchu(dbRows);

        assertEquals(1, records.size());
        OrderRecord record = records.get(0);
        assertEquals("JR001", record.getReqNo());
        assertTrue(record.getStatus().contains("照合中"));
        assertEquals("テスト", record.getUser());
        assertEquals(dbRow, record.getDbValues());
    }
}
