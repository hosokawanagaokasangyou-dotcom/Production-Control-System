package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

import java.nio.file.Files;
import java.nio.file.Path;

import java.util.Map;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledOnOs;
import org.junit.jupiter.api.condition.OS;

import jp.co.pm.ai.desktop.config.AppPaths;

class RequestFormPreviewPdfFontsTest {

    @Test
    void rendererSpec_reflectsCjkScaleFromUi() {
        RequestFormSheetPreviewPdfRenderer.applyCjkMetricsScaleFromUi(Map.of("PM_AI_REQUEST_FORM_PREVIEW_PDF_CJK_SCALE", "0.68"));
        assertEquals("pdfbox-v5-unformatted-date-s68", RequestFormSheetPreviewPdfRenderer.rendererSpec());
        RequestFormSheetPreviewPdfRenderer.applyCjkMetricsScaleFromUi(Map.of());
        assertEquals(
                "pdfbox-v5-unformatted-date-s"
                        + Math.round(AppPaths.DEFAULT_PM_AI_REQUEST_FORM_PREVIEW_PDF_CJK_SCALE * 100),
                RequestFormSheetPreviewPdfRenderer.rendererSpec());
    }

    @Test
    @EnabledOnOs(OS.WINDOWS)
    void load_windowsMsgothicTtc_doesNotThrowHeadMandatory() throws Exception {
        String windir = System.getenv("WINDIR");
        assumeTrue(windir != null && !windir.isBlank());
        Path ttc = Path.of(windir, "Fonts", "msgothic.ttc");
        assumeTrue(Files.isRegularFile(ttc), "msgothic.ttc が無いためスキップ");

        try (PDDocument document = new PDDocument()) {
            RequestFormPreviewPdfFonts.FontPair fonts =
                    RequestFormPreviewPdfFonts.load(document);
            assertNotNull(fonts.regular());
            assertNotNull(fonts.bold());
        }
    }
}
