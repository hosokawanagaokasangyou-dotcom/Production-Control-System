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
        assertTrue(label.contains("1300×250 | 白 | EC,梱包"));
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
}
