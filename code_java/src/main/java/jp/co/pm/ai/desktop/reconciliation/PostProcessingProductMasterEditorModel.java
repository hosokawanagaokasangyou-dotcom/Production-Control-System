package jp.co.pm.ai.desktop.reconciliation;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * 後加工商品マスタ1行分の編集モデル（列名→値）。
 */
public final class PostProcessingProductMasterEditorModel {

    public record ValidationResult(boolean ok, List<String> messages) {
        public static ValidationResult success() {
            return new ValidationResult(true, List.of());
        }

        public static ValidationResult failure(List<String> messages) {
            return new ValidationResult(false, List.copyOf(messages));
        }
    }

    private final List<String> headers;
    private final LinkedHashMap<String, String> values = new LinkedHashMap<>();

    public PostProcessingProductMasterEditorModel(List<String> headers) {
        this.headers =
                List.copyOf(
                        PostProcessingProductMasterColumnGroups.alignHeadersToReference(
                                headers != null ? headers : List.of()));
        for (String h : this.headers) {
            values.put(h, "");
        }
    }

    public List<String> headers() {
        return headers;
    }

    public Map<String, String> snapshot() {
        LinkedHashMap<String, String> copy = new LinkedHashMap<>();
        for (String h : headers) {
            copy.put(h, values.getOrDefault(h, ""));
        }
        return Map.copyOf(copy);
    }

    public String get(String column) {
        return values.getOrDefault(column, "");
    }

    public void set(String column, String value) {
        if (column == null || column.isBlank()) {
            return;
        }
        if (headers.contains(column)) {
            values.put(column, value != null ? value.trim() : "");
        }
    }

    public void loadFromRow(Map<String, String> row) {
        if (row == null) {
            return;
        }
        for (String h : headers) {
            values.put(h, row.getOrDefault(h, "").trim());
        }
    }

    public void applyTemplateRow(Map<String, String> templateRow) {
        loadFromRow(templateRow);
    }

    public ValidationResult validateForUpload(List<String> existingCodesInUpload) {
        List<String> msgs = new ArrayList<>();
        String code = get("商品コード");
        if (code.isBlank()) {
            msgs.add("商品コードを入力してください。");
        }
        if (existingCodesInUpload != null) {
            long dup =
                    existingCodesInUpload.stream()
                            .filter(c -> Objects.equals(code, c != null ? c.trim() : ""))
                            .count();
            if (dup > 1) {
                msgs.add("アップロード用ファイル内で商品コードが重複しています: " + code);
            }
        }
        return msgs.isEmpty() ? ValidationResult.success() : ValidationResult.failure(msgs);
    }
}
