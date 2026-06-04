package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.io.PostProcessingProductMasterIo;

class PostProcessingProductMasterDuplicateCheckTest {

    @AfterEach
    void clearCache() {
        PostProcessingProductMasterReferenceCache.invalidate();
    }

    @Test
    void emptyCodeIsNotUsable() throws Exception {
        PostProcessingProductMasterDuplicateCheck.Result result =
                PostProcessingProductMasterDuplicateCheck.check(
                        "  ", Path.of("missing.xlsx"), List.of(), null);
        assertFalse(result.usable());
        assertTrue(result.messages().get(0).contains("商品コード"));
    }

    @Test
    void detectsDuplicateInUploadRows(@TempDir Path temp) throws Exception {
        List<String> headers = List.of("商品コード", "商品名1");
        Path ref = temp.resolve("ref.xlsx");
        PostProcessingProductMasterIo.writeUploadWorkbook(ref, headers, List.of());

        Map<String, String> existing = new LinkedHashMap<>();
        existing.put("商品コード", "DUP001");
        existing.put("商品名1", "existing");

        PostProcessingProductMasterDuplicateCheck.Result result =
                PostProcessingProductMasterDuplicateCheck.check(
                        "DUP001", ref, List.of(existing), null);
        assertFalse(result.usable());
        assertTrue(result.inUploadRows());
        assertTrue(result.messages().stream().anyMatch(m -> m.contains("アップロード用一覧")));
    }

    @Test
    void excludesSelectedUploadRow(@TempDir Path temp) throws Exception {
        List<String> headers = List.of("商品コード", "商品名1");
        Path ref = temp.resolve("ref.xlsx");
        PostProcessingProductMasterIo.writeUploadWorkbook(ref, headers, List.of());

        Map<String, String> editing = new LinkedHashMap<>();
        editing.put("商品コード", "SELF001");
        editing.put("商品名1", "self");

        PostProcessingProductMasterDuplicateCheck.Result result =
                PostProcessingProductMasterDuplicateCheck.check(
                        "SELF001", ref, List.of(editing), editing);
        assertTrue(result.usable());
        assertFalse(result.inUploadRows());
    }

    @Test
    void detectsDuplicateInReferenceMaster(@TempDir Path temp) throws Exception {
        List<String> headers = List.of("商品コード", "商品名1");
        Path ref = temp.resolve("ref.xlsx");
        Map<String, String> refRow = new LinkedHashMap<>();
        refRow.put("商品コード", "REF001");
        refRow.put("商品名1", "master");
        PostProcessingProductMasterIo.writeUploadWorkbook(ref, headers, List.of(refRow));

        PostProcessingProductMasterDuplicateCheck.Result result =
                PostProcessingProductMasterDuplicateCheck.check(
                        "REF001", ref, List.of(), null);
        assertFalse(result.usable());
        assertTrue(result.inReferenceMaster());
        assertTrue(result.messages().stream().anyMatch(m -> m.contains("参照マスタ")));
    }

    @Test
    void uniqueCodeIsUsable(@TempDir Path temp) throws Exception {
        List<String> headers = List.of("商品コード", "商品名1");
        Path ref = temp.resolve("ref.xlsx");
        Map<String, String> refRow = new LinkedHashMap<>();
        refRow.put("商品コード", "REF001");
        refRow.put("商品名1", "master");
        PostProcessingProductMasterIo.writeUploadWorkbook(ref, headers, List.of(refRow));

        PostProcessingProductMasterDuplicateCheck.Result result =
                PostProcessingProductMasterDuplicateCheck.check(
                        "NEW999", ref, List.of(), null);
        assertTrue(result.usable());
        assertFalse(result.inReferenceMaster());
        assertFalse(result.inUploadRows());
    }
}
