package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.config.AppPaths;

/** 実環境の GB.pdf テキスト形状を確認する（到達不可ならスキップ）。 */
class RequestFormTpiPdfExtractorGbPdfTextProbeTest {

    @Test
    void probeGbPdfTextLayout() throws Exception {
        Path pdf = resolveGbPdf();
        if (pdf == null || !Files.isRegularFile(pdf)) {
            return;
        }
        String text = RequestFormTpiPdfExtractor.readPdfText(pdf.toFile());
        assertFalse(text.isBlank(), "text sample=" + text.substring(0, Math.min(200, text.length())));
        List<Map<String, String>> entries =
                RequestFormTpiPdfExtractor.extractEntries(pdf.toFile(), Map.of());
        assertFalse(entries.isEmpty());
        Map<String, String> raw =
                entries.stream()
                        .filter(e -> "GB60604".equals(e.get("依頼Ｎｏ")))
                        .findFirst()
                        .orElse(entries.get(0));
        assertNotNull(raw.get("依頼Ｎｏ"));
        assertFalse(raw.get("製品").isBlank(), "raw=" + raw);
        assertFalse(raw.get("数量1").isBlank(), "raw=" + raw);
        Map<String, String> raw606 =
                entries.stream()
                        .filter(e -> "GB60606".equals(e.get("依頼Ｎｏ")))
                        .findFirst()
                        .orElse(Map.of());
        if (!raw606.isEmpty()) {
            assertEquals(
                    "15025-NR28-1560X50\n15025-NR28-1560X150",
                    raw606.get("製品"),
                    "raw606=" + raw606);
            assertEquals("15025-NR28-1560X200", raw606.get("原反"), "raw606=" + raw606);
        }
        Map<String, String> raw604 =
                entries.stream()
                        .filter(e -> "GB60604".equals(e.get("依頼Ｎｏ")))
                        .findFirst()
                        .orElse(Map.of());
        if (!raw604.isEmpty()) {
            assertEquals("HB3000GB-45-1050X100", raw604.get("製品"), "raw604=" + raw604);
            assertEquals("HB3000GB-90-1050X100", raw604.get("原反"), "raw604=" + raw604);
        }
        System.out.println("GB.pdf text sample:\n" + text.substring(0, Math.min(500, text.length())));
        System.out.println("raw=" + raw);
    }

    private static Path resolveGbPdf() {
        String uncDir = System.getenv("PM_AI_REQUEST_FORM_TPI_PDF_DIR");
        if (uncDir != null && !uncDir.isBlank()) {
            Path p = Path.of(uncDir, "GB.pdf");
            if (Files.isRegularFile(p)) {
                return p;
            }
        }
        Path factory = Path.of(AppPaths.DEFAULT_PM_AI_REQUEST_FORM_TPI_PDF_DIR_KONAN, "GB.pdf");
        if (Files.isRegularFile(factory)) {
            return factory;
        }
        return null;
    }
}
