package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.Test;

class RequestFormMasterProductCandidateMatcherTest {

    @Test
    void formatCandidateLabel_includesKakoNaiyo() {
        ProductInfo p =
                new ProductInfo(
                        "A2F20AXD0250FN1",
                        "",
                        "15020-NP17",
                        "",
                        "",
                        "",
                        "",
                        "6783",
                        "15020",
                        "1300",
                        "250",
                        "白",
                        "",
                        "EC,梱包");

        String label = RequestFormMasterProductCandidateMatcher.formatCandidateLabel(p);

        assertTrue(label.endsWith(" | EC,梱包"));
        assertTrue(label.contains(" | NP17 | 1300×250 | 白 | EC,梱包"));
    }

    @Test
    void formatTypeForLabel_splitsShohinName1() {
        assertEquals("NP17", RequestFormMasterProductCandidateMatcher.formatTypeForLabel("15020-NP17"));
        assertEquals("WHOLE", RequestFormMasterProductCandidateMatcher.formatTypeForLabel("WHOLE"));
        assertEquals("?", RequestFormMasterProductCandidateMatcher.formatTypeForLabel(""));
        assertEquals("?", RequestFormMasterProductCandidateMatcher.formatTypeForLabel(null));
    }

    @Test
    void formatCandidateLabel_zeroFoamColor_showsDash() {
        ProductInfo p =
                new ProductInfo(
                        "X", "", "", "", "", "", "", "", "", "1", "1", "0", "", "");
        String label = RequestFormMasterProductCandidateMatcher.formatCandidateLabel(p);
        assertTrue(label.contains("1×1 | - | "));
    }

    @Test
    void formatCandidateLabel_emptyFoamColor_showsPlaceholder() {
        ProductInfo p =
                new ProductInfo(
                        "X", "", "", "", "", "", "", "", "", "1", "1", "", "", "");
        String label = RequestFormMasterProductCandidateMatcher.formatCandidateLabel(p);
        assertTrue(label.contains("1×1 | ? | "));
    }

    @Test
    void buildRankedCandidateLabels_similarHinmei_doesNotCrossMatch() {
        List<ProductInfo> catalog =
                List.of(
                        new ProductInfo(
                                "CODE-A",
                                "S1",
                                "15020-NP17",
                                "",
                                "",
                                "",
                                "",
                                "6798",
                                "15020",
                                "1300",
                                "250",
                                "",
                                "",
                                "EC,梱包"),
                        new ProductInfo(
                                "CODE-B",
                                "S2",
                                "15021-NP18",
                                "",
                                "",
                                "",
                                "",
                                "6783",
                                "15021",
                                "1300",
                                "250",
                                "",
                                "",
                                "EC"));

        List<String> labels =
                RequestFormMasterProductCandidateMatcher.buildRankedCandidateLabels(
                        catalog, "CODE-A", "15020", "NP17", "250", "6783", 10);

        assertEquals(1, labels.size());
        assertTrue(labels.get(0).contains("CODE-B"));
    }

    @Test
    void buildRankedCandidateLabels_typeMatch_ranksAbovePartOnlyWithTypeWeight() {
        List<ProductInfo> catalog =
                List.of(
                        new ProductInfo(
                                "CODE-PART-ONLY",
                                "",
                                "15020-OTHER",
                                "",
                                "",
                                "",
                                "",
                                "6783",
                                "15020",
                                "1300",
                                "250",
                                "",
                                "",
                                ""),
                        new ProductInfo(
                                "CODE-TYPE-MATCH",
                                "",
                                "15020-NP17",
                                "",
                                "",
                                "",
                                "",
                                "6783",
                                "15020",
                                "1300",
                                "250",
                                "",
                                "",
                                ""));

        List<String> labels =
                RequestFormMasterProductCandidateMatcher.buildRankedCandidateLabels(
                        catalog, "", "15020", "NP17", "250", "6783", 10);

        assertEquals(2, labels.size());
        assertTrue(labels.get(0).contains("CODE-TYPE-MATCH"));
    }

    @Test
    void buildRankedCandidateLabels_lowercaseKeywords_matchUppercaseMasterFields() {
        List<ProductInfo> catalog =
                List.of(
                        new ProductInfo(
                                "A2F20AXD0250FN1",
                                "S1",
                                "15020-NP17",
                                "",
                                "",
                                "",
                                "",
                                "6783",
                                "15020",
                                "1300",
                                "250",
                                "白",
                                "",
                                "EC,梱包"));

        List<String> labels =
                RequestFormMasterProductCandidateMatcher.buildRankedCandidateLabels(
                        catalog, "a2f20axd0250fn1", "15020", "np17", "250", "6783", 10);

        assertEquals(1, labels.size());
        assertTrue(labels.get(0).contains("A2F20AXD0250FN1"));
    }

    @Test
    void buildRankedCandidateLabels_exactHinmei_returnsOnlyExactFoamName() {
        List<ProductInfo> catalog =
                List.of(
                        new ProductInfo(
                                "X1", "", "", "", "", "", "", "6798", "1", "1", "1", "", "", ""),
                        new ProductInfo(
                                "X2", "", "", "", "", "", "", "6783", "2", "1", "1", "", "", ""));

        List<String> labels =
                RequestFormMasterProductCandidateMatcher.buildRankedCandidateLabels(
                        catalog, "", "", "", "", "6783", 5);

        assertEquals(1, labels.size());
        assertTrue(labels.get(0).contains("X2"));
    }

    @Test
    void filterCatalogByShohinCodePrefixes_matchesAnyPrefix() {
        List<ProductInfo> catalog =
                List.of(
                        new ProductInfo("A2F20AXD0250FN1", "", "", "", "", "", "", "", "", "", "", "", "", ""),
                        new ProductInfo("B1TEST001", "", "", "", "", "", "", "", "", "", "", "", "", ""),
                        new ProductInfo("C9OTHER", "", "", "", "", "", "", "", "", "", "", "", "", ""));

        List<ProductInfo> filtered =
                RequestFormMasterProductCandidateMatcher.filterCatalogByShohinCodePrefixes(
                        catalog, List.of("A2", "B1"));

        assertEquals(2, filtered.size());
        assertEquals("A2F20AXD0250FN1", filtered.get(0).getShohinCode());
        assertEquals("B1TEST001", filtered.get(1).getShohinCode());
    }

    @Test
    void filterCatalogByShohinCodePrefixes_emptyPrefixes_returnsAll() {
        List<ProductInfo> catalog =
                List.of(new ProductInfo("X", "", "", "", "", "", "", "", "", "", "", "", "", ""));

        assertEquals(
                1,
                RequestFormMasterProductCandidateMatcher.filterCatalogByShohinCodePrefixes(
                                catalog, List.of())
                        .size());
    }

    @Test
    void filterCatalogForMasterReferenceSearch_bothSidesConfigured_usesUnion() {
        List<ProductInfo> catalog =
                List.of(
                        new ProductInfo("A2CODE", "", "", "", "", "", "", "", "", "", "", "", "", ""),
                        new ProductInfo("G1RAW", "", "", "", "", "", "", "", "", "", "", "", "", ""),
                        new ProductInfo("Z9SKIP", "", "", "", "", "", "", "", "", "", "", "", "", ""));

        List<ProductInfo> filtered =
                RequestFormMasterProductCandidateMatcher.filterCatalogForMasterReferenceSearch(
                        catalog, List.of("A2"), List.of("G1"));

        assertEquals(2, filtered.size());
        assertEquals("A2CODE", filtered.get(0).getShohinCode());
        assertEquals("G1RAW", filtered.get(1).getShohinCode());
    }

    @Test
    void buildRankedCandidateLabels_respectsPrefixFilterOnCatalog() {
        List<ProductInfo> catalog =
                List.of(
                        new ProductInfo(
                                "A2CODE",
                                "",
                                "15020-NP17",
                                "",
                                "",
                                "",
                                "",
                                "6783",
                                "15020",
                                "1300",
                                "250",
                                "",
                                "",
                                ""),
                        new ProductInfo(
                                "Z9CODE",
                                "",
                                "15021-NP18",
                                "",
                                "",
                                "",
                                "",
                                "6783",
                                "15021",
                                "1300",
                                "250",
                                "",
                                "",
                                ""));

        List<ProductInfo> filtered =
                RequestFormMasterProductCandidateMatcher.filterCatalogByShohinCodePrefixes(
                        catalog, List.of("A2"));

        List<String> labels =
                RequestFormMasterProductCandidateMatcher.buildRankedCandidateLabels(
                        filtered, "", "15020", "NP17", "250", "6783", 10);

        assertEquals(1, labels.size());
        assertTrue(labels.get(0).contains("A2CODE"));
    }
}
