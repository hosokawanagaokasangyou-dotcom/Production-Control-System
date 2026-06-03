package jp.co.pm.ai.desktop.reconciliation;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import javafx.scene.control.ComboBox;
import javafx.scene.control.TextField;

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

    /**
     * 依頼書製品行の内容でマスタ列を上書きする。商品コードは雛形のまま（呼び出し側で雛形適用後に呼ぶ）。
     */
    public void applyRequestFormProductRow(
            ReconciliationApp.ProductRow row, String formKakoKbnLabel) {
        if (row == null) {
            return;
        }
        setIfPresent("発泡体品名", textOf(row.txtHinmei));
        setIfPresent("発泡体品番", textOf(row.txtPart));
        setIfPresent("発泡体タイプ", textOf(row.txtType));
        setIfPresent("発泡体幅", textOf(row.txtWidth));
        setIfPresent("発泡体長さ", textOf(row.txtLength));
        setIfPresent("発泡体色", textOf(row.txtColor));
        setIfPresent("発泡体区分", textOf(row.txtCategory));
        setIfPresent("EC面指定コード", comboValue(row.cmbEcSide));
        setIfPresent("トリミング", mapTrimming(comboValue(row.cmbTrimming)));
        setIfPresent("自社後加工区分", mapSelfKakoKbn(formKakoKbnLabel));
        String seihin = textOf(row.txtSeihinmei);
        if (!seihin.isEmpty()) {
            setIfPresent("製品コード", seihin);
            setIfPresent("商品名2", seihin);
        }
        String item = textOf(row.txtItem);
        if (!item.isEmpty()) {
            setIfPresent("商品コード", item);
        }
        rebuildShohinName1FromFoamSpec();
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

    private void rebuildShohinName1FromFoamSpec() {
        String part = get("発泡体品番");
        String type = get("発泡体タイプ");
        String width = get("発泡体幅");
        String length = get("発泡体長さ");
        String pack = get("発泡体梱等");
        if (part.isEmpty() && type.isEmpty() && width.isEmpty() && length.isEmpty()) {
            return;
        }
        StringBuilder sb = new StringBuilder();
        if (!part.isEmpty()) {
            sb.append(part);
        }
        if (!type.isEmpty()) {
            if (!sb.isEmpty()) {
                sb.append('-');
            }
            sb.append(type);
        }
        if (!width.isEmpty() || !length.isEmpty()) {
            if (!sb.isEmpty()) {
                sb.append('-');
            }
            sb.append(width);
            if (!length.isEmpty()) {
                sb.append('X').append(length);
            }
        }
        if (!pack.isEmpty()) {
            if (!sb.isEmpty()) {
                sb.append('-');
            }
            sb.append(pack);
        }
        if (!sb.isEmpty()) {
            set("商品名1", sb.toString());
        }
    }

    private static void setIfPresent(
            PostProcessingProductMasterEditorModel model, String col, String val) {
        if (val != null && !val.isBlank()) {
            model.set(col, val);
        }
    }

    private void setIfPresent(String col, String val) {
        setIfPresent(this, col, val);
    }

    private static String textOf(TextField field) {
        return field != null && field.getText() != null ? field.getText().trim() : "";
    }

    private static String comboValue(ComboBox<String> combo) {
        return combo != null && combo.getValue() != null ? combo.getValue().trim() : "";
    }

    /** 依頼書コンボ「あり/なし」等をマスタの 0/1 に寄せる。 */
    static String mapTrimming(String label) {
        if (label == null || label.isBlank()) {
            return "";
        }
        String t = label.trim();
        if (t.contains("あり") || t.startsWith("1")) {
            return "1";
        }
        if (t.contains("なし") || t.startsWith("0")) {
            return "0";
        }
        return t;
    }

    static String mapSelfKakoKbn(String formLabel) {
        if (formLabel == null || formLabel.isBlank()) {
            return "";
        }
        if ("後加工".equals(formLabel.trim())) {
            return "1";
        }
        if ("TPI".equals(formLabel.trim())) {
            return "0";
        }
        return formLabel.trim();
    }
}
