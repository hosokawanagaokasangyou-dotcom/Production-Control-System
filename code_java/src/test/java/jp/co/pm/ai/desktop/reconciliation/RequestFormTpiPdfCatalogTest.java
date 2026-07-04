package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import org.junit.jupiter.api.Test;

class RequestFormTpiPdfCatalogTest {

    @Test
    void findForIraiNo_shortStemDoesNotPrefixMatch() throws Exception {
        Path dir = Files.createTempDirectory("tpi-pdf-catalog");
        Path pdf = dir.resolve("GB.pdf");
        Files.write(pdf, new byte[] {0x25, 0x50, 0x44, 0x46});
        Optional<java.io.File> found =
                RequestFormTpiPdfCatalog.findForIraiNo("GB60604", dir.toString());
        assertFalse(found.isPresent(), "GB.pdf must not prefix-match GB60604");
    }

    @Test
    void findForIraiNo_exactFileName() throws Exception {
        Path dir = Files.createTempDirectory("tpi-pdf-catalog-exact");
        Path pdf = dir.resolve("PN04-03.pdf");
        Files.write(pdf, new byte[] {0x25, 0x50, 0x44, 0x46});
        Optional<java.io.File> found =
                RequestFormTpiPdfCatalog.findForIraiNo("PN04-03", dir.toString());
        assertTrue(found.isPresent());
        assertEquals("PN04-03.pdf", found.get().getName());
    }

    @Test
    void findForIraiNo_exactGbStem() throws Exception {
        Path dir = Files.createTempDirectory("tpi-pdf-catalog-gb");
        Path pdf = dir.resolve("GB60604.pdf");
        Files.write(pdf, new byte[] {0x25, 0x50, 0x44, 0x46});
        Optional<java.io.File> found =
                RequestFormTpiPdfCatalog.findForIraiNo("GB60604", dir.toString());
        assertTrue(found.isPresent());
        assertEquals("GB60604.pdf", found.get().getName());
    }

    @Test
    void findForIraiNo_fixtureGbScanned() {
        Path pdf = Path.of("src/test/resources/tpi-request-forms/GB-scanned.pdf");
        if (!Files.isRegularFile(pdf)) {
            return;
        }
        assertFalse(
                RequestFormTpiPdfCatalog.findForIraiNo(
                                "GB60604", pdf.getParent().toString())
                        .isPresent(),
                "GB-scanned.pdf must not filename-prefix-match");
    }

    @Test
    void shouldAutoAddScannedEntry_skipsBundledWhenDedicatedExists() throws Exception {
        Path dir = Files.createTempDirectory("tpi-bundled");
        Path bundle = dir.resolve("GB.pdf");
        Path dedicated = dir.resolve("GB60604.pdf");
        Files.write(bundle, new byte[] {0x25, 0x50, 0x44, 0x46});
        Files.write(dedicated, new byte[] {0x25, 0x50, 0x44, 0x46});
        List<Map<String, String>> entries =
                List.of(
                        Map.of("依頼Ｎｏ", "GB60604"),
                        Map.of("依頼Ｎｏ", "GB60606"),
                        Map.of("依頼Ｎｏ", "GB60605"));
        assertFalse(
                RequestFormTpiPdfCatalog.shouldAutoAddScannedEntry(
                        dir.toFile(), bundle.toFile(), entries, entries.get(0)));
        assertTrue(
                RequestFormTpiPdfCatalog.shouldAutoAddScannedEntry(
                        dir.toFile(), bundle.toFile(), entries, entries.get(1)));
        assertTrue(
                RequestFormTpiPdfCatalog.shouldAutoAddScannedEntry(
                        dir.toFile(), bundle.toFile(), entries, entries.get(2)));
        assertTrue(
                RequestFormTpiPdfCatalog.shouldAutoAddScannedEntry(
                        dir.toFile(), dedicated.toFile(), List.of(entries.get(0)), entries.get(0)));
    }

    @Test
    void canLinkIraiInSharedPdf_linksAllWhenNoDedicated() {
        List<Map<String, String>> entries =
                List.of(Map.of("依頼Ｎｏ", "GB60604"), Map.of("依頼Ｎｏ", "GB60606"));
        File bundle = new File("GB.pdf");
        assertTrue(
                RequestFormTpiPdfCatalog.canLinkIraiInSharedPdf(
                        "GB60604", bundle, entries, null));
        assertTrue(
                RequestFormTpiPdfCatalog.canLinkIraiInSharedPdf(
                        "GB60606", bundle, entries, null));
    }
}
