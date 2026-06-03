package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.LinkedHashMap;
import java.util.Map;
import org.junit.jupiter.api.Test;

class ReconciliationAppJuchuMergeTest {

    @Test
    void mergeContract_fillsOnlyWhenFormContractBlank() {
        Map<String, String> db = new LinkedHashMap<>();
        db.put("契約Ｎｏ", "");
        Map<String, String> raw = Map.of("契約Ｎｏ", "RAW-001");

        ReconciliationApp.mergeJuchuContractNoFromRawWhenBlankOrDifferent(db, raw);

        assertEquals("RAW-001", db.get("契約Ｎｏ"));
    }

    @Test
    void mergeContract_doesNotOverwriteUserEditedValue() {
        Map<String, String> db = new LinkedHashMap<>();
        db.put("契約Ｎｏ", "USER-EDIT");
        Map<String, String> raw = Map.of("契約Ｎｏ", "RAW-001");

        ReconciliationApp.mergeJuchuContractNoFromRawWhenBlankOrDifferent(db, raw);

        assertEquals("USER-EDIT", db.get("契約Ｎｏ"));
    }
}
