package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.config.AppPaths;

class RequestFormTpiPdfOcrIntegrationTest {

    @Test
    void extractScannedGbPdf_whenTesseractAvailable() throws Exception {
        if (!RequestFormTpiPdfOcrReader.isAvailable(Map.of())) {
            return;
        }
        Path pdf = Path.of("src/test/resources/tpi-request-forms/GB-scanned.pdf");
        if (!Files.isRegularFile(pdf)) {
            return;
        }
        assumeTrue(
                RequestFormTpiPdfContentDetector.detect(pdf.toFile())
                        == RequestFormTpiPdfContentKind.IMAGE_SCAN,
                "OCR対象の画像専用PDFではないためスキップ");
        List<Map<String, String>> entries =
                RequestFormTpiPdfExtractor.extractEntries(pdf.toFile(), Map.of());
        assertEquals(1, entries.size());
        Map<String, String> raw = entries.get(0);
        assertEquals(
                RequestFormTpiPdfFieldLayout.READ_MODE_OCR,
                raw.get(RequestFormTpiPdfFieldLayout.META_READ_MODE));
        assertFalse(raw.get("依頼Ｎｏ").isBlank(), "raw=" + raw);
        assertNotNull(raw.get("契約Ｎｏ"));
    }

    @Test
    void textLayerFixture_doesNotRequireTesseractWhenNotInstalled() throws Exception {
        if (AppPaths.resolveTesseractConfig(Map.of()).isPresent()) {
            return;
        }
        Path pdf = Path.of("src/test/resources/tpi-request-forms/GB-scanned.pdf");
        if (!Files.isRegularFile(pdf)) {
            return;
        }
        assertEquals(
                RequestFormTpiPdfContentKind.TEXT,
                RequestFormTpiPdfContentDetector.detect(pdf.toFile()));
    }
}
