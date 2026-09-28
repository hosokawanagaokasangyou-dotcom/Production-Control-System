package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class RequestFormTpiPdfReloadTest {

    @Test
    void explicitDataReload_reextractsEvenWhenParseCacheHits() {
        assertEquals(
                RequestFormTpiPdfReload.Action.REEXTRACT,
                RequestFormTpiPdfReload.decide(true, false, true));
        assertEquals(
                RequestFormTpiPdfReload.Action.REEXTRACT,
                RequestFormTpiPdfReload.decide(true, true, true));
    }

    @Test
    void automaticReload_keepsExcelSkipAndParseCache() {
        assertEquals(
                RequestFormTpiPdfReload.Action.SKIP_EXCEL_DUPLICATE,
                RequestFormTpiPdfReload.decide(false, true, true));
        assertEquals(
                RequestFormTpiPdfReload.Action.USE_CACHE,
                RequestFormTpiPdfReload.decide(false, false, true));
        assertEquals(
                RequestFormTpiPdfReload.Action.REEXTRACT,
                RequestFormTpiPdfReload.decide(false, false, false));
    }
}
