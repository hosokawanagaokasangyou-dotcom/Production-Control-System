package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledIf;

class RequestFormTpiPdfTpiIraiNoTest {

    private static final String DIR =
            "\\\\192.168.0.101\\共有フォルダ\\湖南工場\\湖南共有\\002  加工G\\●配台AIシステム\\共有DATA\\TPI依頼書";

    static boolean dirPresent() {
        return Files.isDirectory(Path.of(DIR));
    }

    @Test
    void parseIraiNoFromFileName_tpiTokenInLongName() {
        assertEquals(
                "TPI07-01",
                RequestFormTpiPdfFieldLayout.parseIraiNoFromFileName("後加工注文書  TPI07-01.pdf"));
        assertEquals(
                "TPI07-01",
                RequestFormTpiPdfFieldLayout.parseIraiNoFromFileName("TPI07-01.pdf"));
    }

    @Test
    void parseIraiNoFromText_tpiWithSpaces() {
        assertEquals(
                "TPI07-01",
                RequestFormTpiPdfFieldLayout.parseIraiNoFromText(
                        "依頼NO. 希望納期\nTPI 07-01 2026 7 22 湖南"));
    }

    @Test
    void parseIraiNoFromFileName_a91StillWorks() {
        assertEquals("A9-1", RequestFormTpiPdfFieldLayout.parseIraiNoFromFileName("A9-1.pdf"));
    }

    @Test
    @EnabledIf("dirPresent")
    void extractFromUncTpi07_hasIraiNo() throws Exception {
        Path found;
        try (var stream = Files.list(Path.of(DIR))) {
            found =
                    stream.filter(p -> p.getFileName().toString().contains("TPI07-01"))
                            .findFirst()
                            .orElseThrow();
        }
        List<Map<String, String>> entries =
                RequestFormTpiPdfExtractor.extractEntries(found.toFile());
        assertEquals(1, entries.size());
        assertEquals("TPI07-01", entries.get(0).get("依頼Ｎｏ"));
        assertTrue(
                RequestFormTpiPdfFieldLayout.META_SOURCE_KIND_TPI_PDF.equals(
                        entries.get(0).get(RequestFormTpiPdfFieldLayout.META_SOURCE_KIND)));
    }
}
