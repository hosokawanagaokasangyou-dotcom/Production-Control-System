package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.Test;

class RequestFormMasterProductCandidateMatcherTest {

    @Test
    void buildRankedCandidateLabels_fuzzyHinmei_returnsMultipleOrderedBySimilarity() {
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
                                "EC"),
                        new ProductInfo(
                                "CODE-C",
                                "S3",
                                "99999-XX",
                                "",
                                "",
                                "",
                                "",
                                "1111",
                                "99999",
                                "1000",
                                "100",
                                "",
                                "",
                                ""));

        List<String> labels =
                RequestFormMasterProductCandidateMatcher.buildRankedCandidateLabels(
                        catalog, "CODE-A", "15020", "NP17", "250", "6783", 10);

        assertTrue(labels.size() >= 2, "品名6783と商品CODE-Aで複数候補を出す");
        assertTrue(labels.get(0).contains("CODE-A"));
        assertTrue(labels.stream().anyMatch(l -> l.contains("CODE-B")));
    }

    @Test
    void buildRankedCandidateLabels_exactHinmei_ranksExactFoamNameFirst() {
        List<ProductInfo> catalog =
                List.of(
                        new ProductInfo(
                                "X1", "", "", "", "", "", "", "6798", "1", "1", "1", "", "", ""),
                        new ProductInfo(
                                "X2", "", "", "", "", "", "", "6783", "2", "1", "1", "", "", ""));

        List<String> labels =
                RequestFormMasterProductCandidateMatcher.buildRankedCandidateLabels(
                        catalog, "", "", "", "", "6783", 5);

        assertEquals(2, labels.size());
        assertTrue(labels.get(0).contains("X2"));
    }
}
