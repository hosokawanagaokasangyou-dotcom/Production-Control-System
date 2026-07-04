package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.font.Standard14Fonts;
import org.junit.jupiter.api.Test;

class RequestFormTpiPdfSplitterTest {

    @Test
    void attachSplitPdfsIfBundled_createsOnePdfPerEntry() throws Exception {
        Path dir = Files.createTempDirectory("tpi-split-test");
        File source = dir.resolve("GB.pdf").toFile();
        File cacheRoot = dir.resolve("preview_cache").toFile();
        try (PDDocument doc = new PDDocument()) {
            String[] iraiNos = {"GB60604", "GB60605", "GB60606"};
            for (String iraiNo : iraiNos) {
                PDPage page = new PDPage();
                doc.addPage(page);
                try (PDPageContentStream cs = new PDPageContentStream(doc, page)) {
                    cs.beginText();
                    cs.setFont(new PDType1Font(Standard14Fonts.FontName.HELVETICA), 12);
                    cs.newLineAtOffset(50, 700);
                    cs.showText("No GB " + iraiNo.substring(2));
                    cs.endText();
                }
            }
            doc.save(source);
        }

        List<Map<String, String>> entries =
                List.of(
                        entry("GB60604"),
                        entry("GB60605"),
                        entry("GB60606"));
        List<Map<String, String>> split =
                RequestFormTpiPdfSplitter.attachSplitPdfsIfBundled(
                        source, entries, cacheRoot, Map.of());

        assertEquals(3, split.size());
        for (Map<String, String> entry : split) {
            String path = entry.get(RequestFormTpiPdfFieldLayout.META_SPLIT_PDF_PATH);
            assertNotNull(path);
            File splitPdf = new File(path);
            assertTrue(splitPdf.isFile(), path);
            assertTrue(splitPdf.length() > 128, path);
        }
    }

    private static Map<String, String> entry(String iraiNo) {
        Map<String, String> raw = new LinkedHashMap<>();
        raw.put("依頼Ｎｏ", iraiNo);
        raw.put("原本ファイル名", "GB.pdf");
        raw.put(RequestFormTpiPdfFieldLayout.META_TPI_LAYOUT, RequestFormTpiPdfFieldLayout.LAYOUT_GB_SLICE);
        return raw;
    }
}
