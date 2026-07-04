package jp.co.pm.ai.desktop.reconciliation;

import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class RequestFormTpiPdfContentDetectorTest {

    @Test
    void detect_textPdfFixture() throws Exception {
        Path pdf = Path.of("src/test/resources/tpi-request-forms/ECOWD-JR260603.pdf");
        assertTrue(Files.isRegularFile(pdf));
        assertEquals(
                RequestFormTpiPdfContentKind.TEXT,
                RequestFormTpiPdfContentDetector.detect(pdf.toFile()));
    }

    @Test
    void detect_scannedPdfFixture() throws Exception {
        Path pdf = Path.of("src/test/resources/tpi-request-forms/GB-scanned.pdf");
        if (!Files.isRegularFile(pdf)) {
            return;
        }
        assertEquals(
                RequestFormTpiPdfContentKind.IMAGE_SCAN,
                RequestFormTpiPdfContentDetector.detect(pdf.toFile()));
    }

    @Test
    void detectFromExtractedText_blankIsImageScan() throws Exception {
        assertEquals(
                RequestFormTpiPdfContentKind.IMAGE_SCAN,
                RequestFormTpiPdfContentDetector.detectFromExtractedText(null, ""));
    }
}
