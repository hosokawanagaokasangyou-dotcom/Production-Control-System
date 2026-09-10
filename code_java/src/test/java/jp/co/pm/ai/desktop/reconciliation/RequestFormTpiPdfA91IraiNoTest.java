package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledIf;

class RequestFormTpiPdfA91IraiNoTest {

    private static final String UNC_A91 =
            "\\\\192.168.0.101\\共有フォルダ\\湖南工場\\湖南共有\\002  加工G\\●配台AIシステム\\共有DATA\\TPI依頼書\\A9-1.pdf";

    static boolean a91Present() {
        return Files.isRegularFile(Path.of(UNC_A91));
    }

    @Test
    void parseIraiNoFromFileName_a91Stem() {
        assertEquals("A9-1", RequestFormTpiPdfFieldLayout.parseIraiNoFromFileName("A9-1.pdf"));
        assertEquals(
                "A9-1", RequestFormTpiPdfFieldLayout.resolveIraiNo("A9-1.pdf", "QR-06-011 unrelated"));
    }

    @Test
    void parseIraiNoFromFileName_doesNotStealGbOrPn() {
        assertEquals("GB60604", RequestFormTpiPdfFieldLayout.parseIraiNoFromFileName("GB60604.pdf"));
        assertEquals(
                "PN04-03",
                RequestFormTpiPdfFieldLayout.parseIraiNoFromFileName("後加工注文書（PN04-03).pdf"));
    }

    @Test
    @EnabledIf("a91Present")
    void extractFromUncA91_hasIraiNo() throws Exception {
        Path pdf = Path.of(UNC_A91);
        List<Map<String, String>> entries = RequestFormTpiPdfExtractor.extractEntries(pdf.toFile());
        assertEquals(1, entries.size());
        assertEquals("A9-1", entries.get(0).get("依頼Ｎｏ"));
        assertTrue(
                RequestFormTpiPdfFieldLayout.META_SOURCE_KIND_TPI_PDF.equals(
                        entries.get(0).get(RequestFormTpiPdfFieldLayout.META_SOURCE_KIND)));
    }
}
