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

    @Test
    void resolveFormActiveValues_tpiPdfPrefersDbValuesOverPdfDefaults() {
        Map<String, String> raw = new LinkedHashMap<>();
        raw.put(
                RequestFormTpiPdfFieldLayout.META_SOURCE_KIND,
                RequestFormTpiPdfFieldLayout.META_SOURCE_KIND_TPI_PDF);
        raw.put("依頼Ｎｏ", "JR001");
        raw.put("品名", "PDF品名");
        Map<String, String> db = Map.of("品名", "受注品名", "ユーザー", "U1");
        OrderRecord record =
                new OrderRecord("JR001", "既存登録 (原本一致・TPI PDF)", "U1", "", "", raw, db);

        Map<String, String> active = ReconciliationApp.resolveFormActiveValues(record);

        assertEquals("受注品名", active.get("品名"));
        assertEquals("U1", active.get("ユーザー"));
    }

    @Test
    void resolveFormActiveValues_tpiPdfWithoutDbUsesPdfDefaults() {
        Map<String, String> raw = new LinkedHashMap<>();
        raw.put(
                RequestFormTpiPdfFieldLayout.META_SOURCE_KIND,
                RequestFormTpiPdfFieldLayout.META_SOURCE_KIND_TPI_PDF);
        raw.put("依頼Ｎｏ", "JR002");
        raw.put("品名", "PDF品名");
        OrderRecord record =
                new OrderRecord("JR002", "新規自動追加 (TPI PDF)", "", "", "", raw, Map.of());

        Map<String, String> active = ReconciliationApp.resolveFormActiveValues(record);

        assertEquals("PDF品名", active.get("品名"));
        assertEquals("TPI", active.get("加工区分"));
    }
}
