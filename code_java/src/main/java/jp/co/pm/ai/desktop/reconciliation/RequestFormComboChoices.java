package jp.co.pm.ai.desktop.reconciliation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 依頼書入力フォームの ComboBox 候補リストと入力欄の既定値。{@link jp.co.pm.ai.desktop.config.DesktopSessionState}
 * とユーザープロファイル経由で永続化する。
 */
public final class RequestFormComboChoices {

    public static final String JSON_KEY = "requestFormComboChoices";
    public static final String JSON_FIELD_DEFAULTS_KEY = "fieldDefaults";

    public static final String KEY_INPUT_KBN = "inputKbn";
    public static final String KEY_KAKO_KBN = "kakoKbn";
    /** @deprecated 入力担当はログイン操作者固定のため永続化しない */
    @Deprecated
    public static final String KEY_INPUT_TANTO = "inputTanto";
    public static final String KEY_WARI_SU = "wariSu";
    public static final String KEY_EC_SIDE = "ecSide";
    public static final String KEY_TRIMMING = "trimming";
    public static final String KEY_FEED_LOC = "feedLoc";
    public static final String KEY_STORAGE_LOC = "storageLoc";
    public static final String KEY_YOTO = "yoto";
    public static final String KEY_USER = "user";
    /** 依頼書フォーム製品行マスタ候補コンボ: 商品コード先頭フィルタ（空なら無制限）。 */
    public static final String KEY_MASTER_CANDIDATE_PREFIX_PRODUCT = "masterCandidatePrefixProduct";
    /** 依頼書フォーム原反行マスタ候補コンボ: 商品コード先頭フィルタ（空なら無制限）。 */
    public static final String KEY_MASTER_CANDIDATE_PREFIX_RAW = "masterCandidatePrefixRaw";

    private static final List<String> ALL_KEYS =
            List.of(
                    KEY_INPUT_KBN,
                    KEY_KAKO_KBN,
                    KEY_WARI_SU,
                    KEY_EC_SIDE,
                    KEY_TRIMMING,
                    KEY_FEED_LOC,
                    KEY_STORAGE_LOC,
                    KEY_YOTO,
                    KEY_USER,
                    KEY_MASTER_CANDIDATE_PREFIX_PRODUCT,
                    KEY_MASTER_CANDIDATE_PREFIX_RAW);

    private static final List<String> FIELD_DEFAULT_KEYS = List.of(KEY_INPUT_KBN, KEY_KAKO_KBN);

    private static final ObjectMapper JSON = new ObjectMapper();

    private final Map<String, List<String>> byKey;
    private final Map<String, String> fieldDefaults;

    private RequestFormComboChoices(
            Map<String, List<String>> byKey, Map<String, String> fieldDefaults) {
        this.byKey = Map.copyOf(byKey);
        this.fieldDefaults =
                Map.copyOf(fieldDefaults != null ? fieldDefaults : Map.of());
    }

    public static RequestFormComboChoices empty() {
        return new RequestFormComboChoices(Map.of(), Map.of());
    }

    public static RequestFormComboChoices of(Map<String, List<String>> source) {
        return of(source, Map.of());
    }

    public static RequestFormComboChoices of(
            Map<String, List<String>> source, Map<String, String> fieldDefaults) {
        LinkedHashMap<String, List<String>> normalized = new LinkedHashMap<>();
        if (source != null) {
            for (String key : ALL_KEYS) {
                List<String> values = sanitizeList(source.get(key));
                if (!values.isEmpty()) {
                    normalized.put(key, values);
                }
            }
        }
        LinkedHashMap<String, String> normalizedDefaults = new LinkedHashMap<>();
        if (fieldDefaults != null) {
            for (String key : FIELD_DEFAULT_KEYS) {
                String value = fieldDefaults.get(key);
                if (value != null) {
                    String text = value.strip();
                    if (!text.isEmpty()) {
                        normalizedDefaults.put(key, text);
                    }
                }
            }
        }
        if (normalized.isEmpty() && normalizedDefaults.isEmpty()) {
            return empty();
        }
        return new RequestFormComboChoices(normalized, normalizedDefaults);
    }

    /** ソースコード既定（未保存・工場出荷時の初期値）。 */
    public static RequestFormComboChoices bundledDefaults() {
        LinkedHashMap<String, List<String>> map = new LinkedHashMap<>();
        map.put(KEY_INPUT_KBN, List.of("通常入力", "例外入力"));
        map.put(KEY_KAKO_KBN, List.of("後加工", "TPI"));
        map.put(KEY_WARI_SU, List.of("1", "2", "3", "5", "6", "7", "8", "9", "10"));
        map.put(
                KEY_EC_SIDE,
                List.of(
                        "Ｈ面",
                        "Ｑ面",
                        "両面",
                        "ｽﾗｲｽ面",
                        "ｽｷﾝ面",
                        "Ｑ面/-",
                        "Ｈ面/-",
                        "Ｑ面/Ｑ面/-",
                        "H面/H面/-"));
        map.put(KEY_TRIMMING, List.of("有", "無", "-"));
        map.put(KEY_FEED_LOC, List.of("EC", "SEC", "ｽﾗｲｽ", "ｽﾘｯﾄ", "ｴﾝﾎﾞｽ", "検反", "融着"));
        map.put(
                KEY_STORAGE_LOC,
                List.of(
                        "滋賀",
                        "湖南",
                        "滋賀/湖南",
                        "湖南/中央",
                        "山田",
                        "中山",
                        "中央湖東",
                        "湖南/滋賀",
                        "奥田"));
        map.put(
                KEY_YOTO,
                List.of(
                        "W（自動車）",
                        "B（輸出）",
                        "Y（工材）",
                        "V（TPI）",
                        "A（TPI）",
                        "JR（屋根）",
                        "P（TPI）",
                        "小口加工"));
        map.put(
                KEY_USER,
                List.of(
                        "自動転記",
                        "ｵｶﾓﾄ",
                        "ﾀﾂﾀ",
                        "共和ﾚｻﾞｰ",
                        "Scientex",
                        "共和興",
                        "ｻｶｲﾅｺﾞﾔ",
                        "ﾀﾞｲｳﾚ",
                        "在ｴﾙ",
                        "U4059",
                        "U5001",
                        "張家港",
                        "ｲｽﾞﾐ",
                        "盟和",
                        "高山産業",
                        "中央物産"));
        return new RequestFormComboChoices(map, bundledFieldDefaultsMap());
    }

    /** 【作業指示】入力区分・加工区分のソース既定。 */
    public static Map<String, String> bundledFieldDefaultsMap() {
        return Map.of(KEY_INPUT_KBN, "通常入力", KEY_KAKO_KBN, "後加工");
    }

    public boolean isEmpty() {
        return byKey.isEmpty() && fieldDefaults.isEmpty();
    }

    public Map<String, List<String>> asMap() {
        return byKey;
    }

    public Map<String, String> fieldDefaultsAsMap() {
        return fieldDefaults;
    }

    /** 未保存キーは {@link #bundledDefaults()} にフォールバック。 */
    public List<String> optionsFor(String key) {
        List<String> saved = byKey.get(key);
        if (saved != null && !saved.isEmpty()) {
            return saved;
        }
        return bundledDefaults().byKey.getOrDefault(key, List.of());
    }

    /** 保存済みの既定値（空なら bundled の文字列。候補リストとの整合は {@link #effectiveDefaultFor} で行う）。 */
    public String defaultFor(String key) {
        String saved = fieldDefaults.get(key);
        if (saved != null && !saved.isBlank()) {
            return saved.strip();
        }
        return bundledFieldDefaultsMap().getOrDefault(key, "");
    }

    /**
     * 新規行・フォームクリア時に使う既定値。保存値が候補に無いときは bundled 既定、それも無ければ候補先頭。
     */
    public String effectiveDefaultFor(String key) {
        List<String> options = optionsFor(key);
        for (String candidate : List.of(defaultFor(key), bundledFieldDefaultsMap().getOrDefault(key, ""))) {
            if (!candidate.isBlank() && options.contains(candidate)) {
                return candidate;
            }
        }
        return options.isEmpty() ? "" : options.get(0);
    }

    /**
     * 欠落キー（未保存・空リスト）は bundled 既定で補完する。保存済みの非空リストはユーザーの完全な意思として採用し、
     * bundled を足し戻さない（設定タブでの削除が再読込で復活しないようにする）。
     */
    public RequestFormComboChoices mergedWithDefaults() {
        LinkedHashMap<String, List<String>> mergedLists = new LinkedHashMap<>();
        RequestFormComboChoices bundled = bundledDefaults();
        for (String key : ALL_KEYS) {
            List<String> saved = sanitizeList(byKey.get(key));
            List<String> bundledOpts = sanitizeList(bundled.byKey.get(key));
            if (saved.isEmpty()) {
                if (!bundledOpts.isEmpty()) {
                    mergedLists.put(key, bundledOpts);
                }
            } else {
                mergedLists.put(key, saved);
            }
        }
        LinkedHashMap<String, String> mergedDefaults = new LinkedHashMap<>(bundledFieldDefaultsMap());
        mergedDefaults.putAll(fieldDefaults);
        return new RequestFormComboChoices(mergedLists, mergedDefaults);
    }

    /** 先頭リストの順序を保ち、続けて bundled 側の未登録値を末尾に足す。 */
    static List<String> unionDistinct(List<String> primary, List<String> secondary) {
        LinkedHashMap<String, String> ordered = new LinkedHashMap<>();
        appendDistinct(ordered, primary);
        appendDistinct(ordered, secondary);
        return List.copyOf(ordered.values());
    }

    private static void appendDistinct(LinkedHashMap<String, String> target, List<String> source) {
        if (source == null || source.isEmpty()) {
            return;
        }
        for (String value : source) {
            if (value == null) {
                continue;
            }
            String text = value.strip();
            if (!text.isEmpty()) {
                target.putIfAbsent(text, text);
            }
        }
    }

    /** {@code session_defaults*.json} 等、{@link #JSON_KEY} 配下に候補がある JSON 根から読む。 */
    public static RequestFormComboChoices fromJson(JsonNode root) {
        if (root == null || !root.isObject()) {
            return empty();
        }
        JsonNode node = root.get(JSON_KEY);
        if (node == null || !node.isObject()) {
            return empty();
        }
        return fromChoicesObject(node);
    }

    /**
     * {@link jp.co.pm.ai.desktop.reconciliation.RequestFormInputSettingsStore} の settings ファイル根から読む。
     * {@link #JSON_KEY} 配下を優先し、無ければ根直下の候補キー（移行用）も見る。
     */
    public static RequestFormComboChoices fromSettingsFileRoot(JsonNode root) {
        if (root == null || !root.isObject()) {
            return empty();
        }
        JsonNode nested = root.get(JSON_KEY);
        if (nested != null && nested.isObject()) {
            RequestFormComboChoices fromNested = fromChoicesObject(nested);
            if (!fromNested.isEmpty()) {
                return fromNested;
            }
        }
        if (hasAnyComboListKey(root)) {
            return fromChoicesObject(root);
        }
        return empty();
    }

    private static boolean hasAnyComboListKey(JsonNode root) {
        for (String key : ALL_KEYS) {
            JsonNode arr = root.get(key);
            if (arr != null && arr.isArray()) {
                return true;
            }
        }
        return false;
    }

    private static RequestFormComboChoices fromChoicesObject(JsonNode node) {
        if (node == null || !node.isObject()) {
            return empty();
        }
        LinkedHashMap<String, List<String>> map = new LinkedHashMap<>();
        for (String key : ALL_KEYS) {
            JsonNode arr = node.get(key);
            if (arr == null || !arr.isArray()) {
                continue;
            }
            List<String> values = new ArrayList<>();
            for (JsonNode el : arr) {
                if (el == null || el.isNull()) {
                    continue;
                }
                if (!el.isTextual() && !el.isNumber()) {
                    continue;
                }
                String text = el.asText("").strip();
                if (!text.isEmpty() && !values.contains(text)) {
                    values.add(text);
                }
            }
            if (!values.isEmpty()) {
                map.put(key, List.copyOf(values));
            }
        }
        LinkedHashMap<String, String> defaults = new LinkedHashMap<>();
        JsonNode defaultsNode = node.get(JSON_FIELD_DEFAULTS_KEY);
        if (defaultsNode != null && defaultsNode.isObject()) {
            for (String key : FIELD_DEFAULT_KEYS) {
                JsonNode el = defaultsNode.get(key);
                if (el != null && el.isTextual()) {
                    String text = el.asText("").strip();
                    if (!text.isEmpty()) {
                        defaults.put(key, text);
                    }
                }
            }
        }
        return of(map, defaults);
    }

    public void writeToObjectNode(ObjectNode root) {
        if (root == null || isEmpty()) {
            return;
        }
        root.remove(JSON_KEY);
        ObjectNode choices = root.putObject(JSON_KEY);
        writeChoicesBody(choices);
    }

    /**
     * 設定 JSON 根へ候補をマージする。スナップショットに無いキー（空リストで落ちた投入場所など）は既存値を残す。
     */
    public void mergeIntoSettingsRoot(ObjectNode root) {
        if (root == null || isEmpty()) {
            return;
        }
        ObjectNode choices;
        JsonNode existing = root.get(JSON_KEY);
        if (existing != null && existing.isObject()) {
            choices = (ObjectNode) existing;
        } else {
            choices = root.putObject(JSON_KEY);
        }
        writeChoicesBody(choices);
    }

    private void writeChoicesBody(ObjectNode choices) {
        for (Map.Entry<String, List<String>> entry : byKey.entrySet()) {
            ArrayNode arr = choices.putArray(entry.getKey());
            for (String value : entry.getValue()) {
                arr.add(value);
            }
        }
        if (!fieldDefaults.isEmpty()) {
            ObjectNode defaultsNode;
            JsonNode existingDefaults = choices.get(JSON_FIELD_DEFAULTS_KEY);
            if (existingDefaults != null && existingDefaults.isObject()) {
                defaultsNode = (ObjectNode) existingDefaults;
            } else {
                defaultsNode = choices.putObject(JSON_FIELD_DEFAULTS_KEY);
            }
            for (Map.Entry<String, String> entry : fieldDefaults.entrySet()) {
                defaultsNode.put(entry.getKey(), entry.getValue());
            }
        }
    }

    private static List<String> sanitizeList(List<String> source) {
        if (source == null || source.isEmpty()) {
            return List.of();
        }
        List<String> out = new ArrayList<>();
        for (String value : source) {
            if (value == null) {
                continue;
            }
            String text = value.strip();
            if (!text.isEmpty() && !out.contains(text)) {
                out.add(text);
            }
        }
        return List.copyOf(out);
    }
}
