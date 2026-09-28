package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class RequestFormTpiPdfReloadTest {

    @Test
    void unchangedFile_usesParseCache() {
        assertEquals(
                RequestFormTpiPdfReload.Action.USE_CACHE,
                RequestFormTpiPdfReload.decide(false, true));
    }

    @Test
    void newOrChangedFile_reextracts() {
        assertEquals(
                RequestFormTpiPdfReload.Action.REEXTRACT,
                RequestFormTpiPdfReload.decide(false, false));
        assertEquals(
                RequestFormTpiPdfReload.Action.REEXTRACT,
                RequestFormTpiPdfReload.decide(true, false));
    }

    @Test
    void unchangedExcelDuplicate_skipsWithoutReocr() {
        assertEquals(
                RequestFormTpiPdfReload.Action.SKIP_EXCEL_DUPLICATE,
                RequestFormTpiPdfReload.decide(true, true));
    }
}
