package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.nio.file.Path;
import java.util.List;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.io.PostProcessingProductMasterIo;

class PostProcessingProductMasterSearchTest {

    private static final Path DUMMY_REF = Path.of("nonexistent-reference.xlsx");

    private static List<ProductInfo> sampleCatalog() {
        return List.of(
                new ProductInfo("A2CODE", "", "", "", "", "", "", "", "", "", "", "", "", ""),
                new ProductInfo("G1RAW", "", "", "", "", "", "", "", "", "", "", "", "", ""),
                new ProductInfo("Z9SKIP", "", "", "", "", "", "", "", "", "", "", "", "", ""));
    }

    @Test
    void searchReference_withBothPrefixSides_filtersUnion() throws Exception {
        List<PostProcessingProductMasterIo.SearchHit> hits =
                PostProcessingProductMasterSearch.searchReference(
                        DUMMY_REF,
                        PostProcessingProductMasterIo.SearchFilter.empty(),
                        50,
                        sampleCatalog(),
                        new PostProcessingProductMasterSearch.MasterReferencePrefixFilters(
                                List.of("A2"), List.of("G1")));

        assertEquals(2, hits.size());
        assertEquals("A2CODE", hits.get(0).shohinCode());
        assertEquals("G1RAW", hits.get(1).shohinCode());
    }

    @Test
    void searchReference_withProductSideOnly_keepsRawSideUnrestricted() throws Exception {
        List<PostProcessingProductMasterIo.SearchHit> hits =
                PostProcessingProductMasterSearch.searchReference(
                        DUMMY_REF,
                        PostProcessingProductMasterIo.SearchFilter.empty(),
                        50,
                        sampleCatalog(),
                        new PostProcessingProductMasterSearch.MasterReferencePrefixFilters(
                                List.of("A2"), List.of()));

        assertEquals(3, hits.size());
    }

    @Test
    void filterCatalogForMasterReferenceSearch_withRawSideOnly_keepsProductSideUnrestricted() {
        List<ProductInfo> filtered =
                RequestFormMasterProductCandidateMatcher.filterCatalogForMasterReferenceSearch(
                        sampleCatalog(), List.of(), List.of("G1"));

        assertEquals(3, filtered.size());
    }
}
