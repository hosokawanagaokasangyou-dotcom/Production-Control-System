package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.Test;

class PostProcessingProductMasterEditorModelTest {

    @Test
    void mapSelfKakoKbn_mapsKnownLabels() {
        assertEquals("1", PostProcessingProductMasterEditorModel.mapSelfKakoKbn("後加工"));
        assertEquals("0", PostProcessingProductMasterEditorModel.mapSelfKakoKbn("TPI"));
    }

    @Test
    void mapTrimming_mapsLabels() {
        assertEquals("1", PostProcessingProductMasterEditorModel.mapTrimming("1:あり"));
        assertEquals("0", PostProcessingProductMasterEditorModel.mapTrimming("0:なし"));
    }

    @Test
    void validateRequiresShohinCode() {
        List<String> headers = List.of("商品コード", "商品名1");
        PostProcessingProductMasterEditorModel model =
                new PostProcessingProductMasterEditorModel(headers);
        var result = model.validateForUpload(List.of());
        assertTrue(!result.ok());
        assertTrue(result.messages().stream().anyMatch(m -> m.contains("商品コード")));
    }

    @Test
    void applyTemplateCopiesValues() {
        List<String> headers = List.of("商品コード", "発泡体品番", "商品名1");
        PostProcessingProductMasterEditorModel model =
                new PostProcessingProductMasterEditorModel(headers);
        model.applyTemplateRow(
                java.util.Map.of(
                        "商品コード", "TEST001",
                        "発泡体品番", "40100",
                        "商品名1", "name"));
        assertEquals("TEST001", model.get("商品コード"));
        assertEquals("40100", model.get("発泡体品番"));
    }
}
