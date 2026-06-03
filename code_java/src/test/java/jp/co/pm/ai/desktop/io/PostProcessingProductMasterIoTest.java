package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.reconciliation.PostProcessingProductMasterColumnGroups;

class PostProcessingProductMasterIoTest {

    @Test
    void roundTripUploadFile(@TempDir Path temp) throws Exception {
        Path ref =
                Path.of("..", "アラジンマスタ", "後加工商品マスタ.xlsx")
                        .toAbsolutePath()
                        .normalize();
        if (!Files.isRegularFile(ref)) {
            List<String> headers =
                    PostProcessingProductMasterColumnGroups.alignHeadersToReference(
                            List.of("商品コード", "商品名1", "発泡体品番"));
            Path upload = temp.resolve("upload.xlsx");
            Map<String, String> row = new LinkedHashMap<>();
            for (String h : headers) {
                row.put(h, "");
            }
            row.put("商品コード", "UNITTEST01");
            row.put("商品名1", "test-name");
            PostProcessingProductMasterIo.writeUploadWorkbook(upload, headers, List.of(row));
            var sheet = PostProcessingProductMasterIo.readUploadWorkbook(upload);
            assertEquals(headers.size(), sheet.headers().size());
            assertEquals(1, sheet.rows().size());
            assertEquals("UNITTEST01", sheet.rows().get(0).get(0));
            return;
        }

        List<String> headers = PostProcessingProductMasterIo.readHeaders(ref);
        assertFalse(headers.isEmpty());
        Path upload = temp.resolve("upload.xlsx");
        PostProcessingProductMasterIo.createEmptyUploadFromReference(ref, upload);
        Map<String, String> row = new LinkedHashMap<>();
        for (String h : headers) {
            row.put(h, "");
        }
        row.put("商品コード", "UNITTEST02");
        row.put("商品名1", "round-trip");
        PostProcessingProductMasterIo.writeUploadWorkbook(upload, headers, List.of(row));
        PostProcessingProductMasterColumnGroups.validateHeadersMatch(
                headers, PostProcessingProductMasterIo.readUploadWorkbook(upload).headers());
        var loaded = PostProcessingProductMasterIo.readUploadWorkbook(upload);
        assertEquals(1, loaded.rows().size());
        Map<String, String> map =
                PostProcessingProductMasterIo.rowToMap(loaded.headers(), loaded.rows().get(0));
        assertEquals("UNITTEST02", map.get("商品コード"));
    }
}
