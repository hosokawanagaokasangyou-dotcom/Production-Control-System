package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * 後加工商品マスタ編集フォームの商品コード重複確認。
 */
public final class PostProcessingProductMasterDuplicateCheck {

    public record Result(
            boolean usable,
            boolean inReferenceMaster,
            boolean inUploadRows,
            String shohinCode,
            List<String> messages) {}

    private PostProcessingProductMasterDuplicateCheck() {}

    /**
     * @param excludeUploadRow アップロード一覧で編集中の行（同一コードの自己一致を除外）
     */
    public static Result check(
            String shohinCode,
            Path referencePath,
            Iterable<Map<String, String>> uploadRows,
            Map<String, String> excludeUploadRow)
            throws IOException {
        String code = shohinCode != null ? shohinCode.trim() : "";
        if (code.isEmpty()) {
            return new Result(
                    false,
                    false,
                    false,
                    code,
                    List.of("商品コードを入力してください。"));
        }

        List<String> messages = new ArrayList<>();
        boolean inReference = false;
        boolean inUpload = false;

        Map<String, String> refRow =
                PostProcessingProductMasterReferenceCache.rowByCode(referencePath, code);
        if (refRow != null && !refRow.isEmpty()) {
            inReference = true;
            messages.add("参照マスタ（後加工商品マスタ）に既に存在します: " + code);
        }

        if (uploadRows != null) {
            for (Map<String, String> row : uploadRows) {
                if (row == excludeUploadRow) {
                    continue;
                }
                String existing = row != null ? row.getOrDefault("商品コード", "").trim() : "";
                if (Objects.equals(code, existing)) {
                    inUpload = true;
                    messages.add("アップロード用一覧に既に存在します: " + code);
                    break;
                }
            }
        }

        if (!inReference && !inUpload) {
            messages.add("商品コード " + code + " は重複していません（使用可能）。");
        }

        return new Result(
                !inReference && !inUpload,
                inReference,
                inUpload,
                code,
                List.copyOf(messages));
    }
}
