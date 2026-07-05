package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.LinkedHashMap;
import java.util.Map;

import org.junit.jupiter.api.Test;

class RequestFormPipelineStatusServiceTpiPdfTest {

    @Test
    void resolveOriginalFileName_prefersSplitPdfName() {
        Map<String, String> raw = new LinkedHashMap<>();
        raw.put(RequestFormTpiPdfFieldLayout.META_SPLIT_PDF_PATH, "C:/cache/GB___GB60606.pdf");
        raw.put("原本ファイル名", "GB.pdf");
        assertEquals("GB___GB60606.pdf", RequestFormPipelineStatusService.resolveOriginalFileName(raw));
    }

    @Test
    void buildOriginalDbFromRaw_usesTpiDefaultsForTpiPdfSource() {
        Map<String, String> raw = new LinkedHashMap<>();
        raw.put(
                RequestFormTpiPdfFieldLayout.META_SOURCE_KIND,
                RequestFormTpiPdfFieldLayout.META_SOURCE_KIND_TPI_PDF);
        raw.put("加工内容", "巻き返し");
        Map<String, String> db = RequestFormPipelineStatusService.buildOriginalDbFromRaw(raw);
        assertEquals("TPI", db.get("加工区分"));
        assertEquals("巻き返し", db.get("加工内容"));
    }

    @Test
    void buildOriginalDbFromRaw_usesExcelDefaultsForWorkbookSource() {
        Map<String, String> raw = Map.of("品名", "NR28");
        Map<String, String> db = RequestFormPipelineStatusService.buildOriginalDbFromRaw(raw);
        assertFalse(db.containsKey("加工区分"));
        assertEquals("NR28", db.get("品名"));
    }

    @Test
    void resolveOriginalFileName_usesTaggedSourceFileName() {
        Map<String, String> raw = Map.of("_sourceFileName", "GB60606.xlsm");
        assertEquals("GB60606.xlsm", RequestFormPipelineStatusService.resolveOriginalFileName(raw));
    }

    @Test
    void resolveOriginalFileName_returnsEmptyForNullRaw() {
        assertEquals("", RequestFormPipelineStatusService.resolveOriginalFileName(null));
    }

    @Test
    void resolveOriginalFileName_fallsBackToOriginalFileNameField() {
        Map<String, String> raw = Map.of("原本ファイル名", "bundle.pdf");
        assertEquals("bundle.pdf", RequestFormPipelineStatusService.resolveOriginalFileName(raw));
    }

    @Test
    void buildOriginalDbFromRaw_tpiFlagIsRequiredForTpiDefaults() {
        Map<String, String> raw = Map.of("加工内容", "巻き返し");
        Map<String, String> db = RequestFormPipelineStatusService.buildOriginalDbFromRaw(raw);
        assertTrue(db.isEmpty() || !"TPI".equals(db.get("加工区分")));
    }
}
