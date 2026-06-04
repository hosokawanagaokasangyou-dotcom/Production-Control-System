package jp.co.pm.ai.desktop.reconciliation;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

/**
 * 後加工商品マスタ「区分情報」「発泡体」タブの区分コード選択肢（Aladdin 画面相当）。
 */
public final class PostProcessingProductMasterKubunChoices {

    public record Option(String code, String label) {}

    private static final Map<String, List<Option>> BY_COLUMN = build();

    private PostProcessingProductMasterKubunChoices() {}

    public static boolean hasChoices(String columnName) {
        return columnName != null && BY_COLUMN.containsKey(columnName.trim());
    }

    public static List<String> pickerLabels(String columnName) {
        List<Option> opts = options(columnName);
        if (opts.isEmpty()) {
            return List.of();
        }
        List<String> labels = new ArrayList<>(opts.size());
        for (Option o : opts) {
            if (o.code().equals(o.label())) {
                labels.add(o.code());
            } else {
                labels.add(displayLabel(o.code(), o.label()));
            }
        }
        return List.copyOf(labels);
    }

    public static String displayLabel(String code, String label) {
        String c = code != null ? code.trim() : "";
        String n = label != null ? label.trim() : "";
        if (c.isEmpty() && n.isEmpty()) {
            return "";
        }
        if (n.isEmpty()) {
            return c;
        }
        if (c.isEmpty()) {
            return n;
        }
        return c + ":" + n;
    }

    public static String resolveLabel(String columnName, String rawCode) {
        String code = normalizeCode(columnName, rawCode);
        if (code.isEmpty()) {
            return "";
        }
        for (Option o : options(columnName)) {
            if (o.code().equals(code)) {
                return o.label();
            }
        }
        return "";
    }

    public static String normalizeCode(String columnName, String raw) {
        if (columnName == null || !hasChoices(columnName)) {
            return raw != null ? raw.trim() : "";
        }
        String val = raw != null ? raw.trim() : "";
        if (val.isEmpty()) {
            return "";
        }
        if (val.endsWith(".0")) {
            val = val.substring(0, val.length() - 2).trim();
        }
        int colon = val.indexOf(':');
        if (colon >= 0) {
            val = val.substring(0, colon).trim();
        }
        try {
            int n = (int) Double.parseDouble(val);
            val = String.valueOf(n);
        } catch (NumberFormatException ignored) {
        }
        for (Option o : options(columnName)) {
            if (o.code().equals(val)) {
                return o.code();
            }
            if (o.label().equals(val)) {
                return o.code();
            }
        }
        return val;
    }

    public static String resolveCodeFromPickerInput(String columnName, String text) {
        if (text == null || text.isBlank()) {
            return "";
        }
        String trimmed = text.trim();
        String norm = normalizeCode(columnName, trimmed);
        for (Option o : options(columnName)) {
            String label = displayLabel(o.code(), o.label());
            if (label.equals(trimmed) || label.startsWith(norm + ":")) {
                return o.code();
            }
        }
        return norm;
    }

    private static List<Option> options(String columnName) {
        if (columnName == null) {
            return List.of();
        }
        return BY_COLUMN.getOrDefault(columnName.trim(), List.of());
    }

    private static Map<String, List<Option>> build() {
        Map<String, List<Option>> m = new LinkedHashMap<>();
        m.put("品区分", List.of(new Option("0", "商品"), new Option("1", "製品"), new Option("2", "部品")));
        m.put(
                "自社後加工区分",
                List.of(new Option("0", "自社加工"), new Option("1", "後加工")));
        m.put(
                "加工単価区分",
                List.of(new Option("0", "積上"), new Option("1", "打換")));
        m.put(
                "展開区分",
                List.of(
                        new Option("0", "しない"),
                        new Option("1", "組立展開"),
                        new Option("2", "売上展開")));
        m.put(
                "ロット管理区分",
                List.of(new Option("0", "対象"), new Option("1", "対象外")));
        m.put(
                "在庫管理区分",
                List.of(new Option("0", "対象"), new Option("1", "対象外")));
        m.put(
                "税率区分コード",
                List.of(new Option("0", "非課税"), new Option("1", "課税")));
        m.put(
                "名称入力区分",
                List.of(new Option("0", "しない"), new Option("1", "する")));
        m.put(
                "削除区分",
                List.of(new Option("0", "しない"), new Option("1", "する")));
        m.put(
                "原価単価取得区分",
                List.of(
                        new Option("1", "マスタ原価"),
                        new Option("2", "在庫単価"),
                        new Option("3", "積上原価"),
                        new Option("4", "原価掛率")));
        m.put(
                "売上時原価引当区分",
                List.of(new Option("0", "しない"), new Option("1", "する")));
        m.put(
                "直送原価取得区分",
                List.of(new Option("1", "通常原価"), new Option("2", "仕入金額")));
        m.put(
                "手配区分",
                List.of(
                        new Option("0", "なし"),
                        new Option("1", "直送"),
                        new Option("2", "手配")));
        m.put(
                "原価洗替区分",
                List.of(new Option("0", "しない"), new Option("1", "する")));
        m.put(
                "AEC連携対象フラグ",
                List.of(new Option("0", "対象外"), new Option("1", "対象")));

        // --- 発泡体タブ（スクリーンショット相当） ---
        m.put(
                "トリミング",
                List.of(new Option("0", "なし"), new Option("1", "あり")));
        m.put(
                "UL規格",
                List.of(new Option("0", "対象外"), new Option("1", "対象")));
        m.put(
                "長さ換算区分",
                List.of(new Option("0", "対象外"), new Option("1", "対象")));
        m.put(
                "融着",
                List.of(new Option("0", "しない"), new Option("1", "する")));

        RequestFormComboChoices bundled = RequestFormComboChoices.bundledDefaults();
        m.put("EC面指定コード", stringOptions(bundled.optionsFor(RequestFormComboChoices.KEY_EC_SIDE)));
        m.put("ユーザ", stringOptions(bundled.optionsFor(RequestFormComboChoices.KEY_USER)));
        m.put(
                "在庫場所",
                stringOptions(bundled.optionsFor(RequestFormComboChoices.KEY_STORAGE_LOC)));

        return Map.copyOf(m);
    }

    private static List<Option> stringOptions(List<String> labels) {
        if (labels == null || labels.isEmpty()) {
            return List.of();
        }
        List<Option> out = new ArrayList<>();
        for (String label : labels) {
            if (label == null || label.isBlank()) {
                continue;
            }
            String t = label.trim();
            out.add(new Option(t, t));
        }
        return List.copyOf(out);
    }

    /** テスト・将来拡張用。 */
    static Optional<List<Option>> optionsForColumn(String columnName) {
        List<Option> opts = options(columnName);
        return opts.isEmpty() ? Optional.empty() : Optional.of(opts);
    }
}
