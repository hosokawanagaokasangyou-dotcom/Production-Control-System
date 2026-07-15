package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.DeserializationFeature;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.ObjectReader;

/** 「担当OP_限定」セルの厳格 JSON 文字列配列 codec。 */
public final class LimitedOperatorJsonCodec {

    private static final ObjectMapper JSON = new ObjectMapper();
    private static final ObjectReader STRICT_JSON_READER =
            JSON.readerFor(JsonNode.class)
                    .with(DeserializationFeature.FAIL_ON_TRAILING_TOKENS);

    private LimitedOperatorJsonCodec() {}

    public static List<String> decode(String cellValue) {
        String raw = cellValue != null ? cellValue.strip() : "";
        if (raw.isEmpty()) {
            return List.of();
        }
        final JsonNode root;
        try {
            root = STRICT_JSON_READER.readValue(raw);
        } catch (JsonProcessingException ex) {
            throw new IllegalArgumentException("担当OP_限定は正しいJSON文字列配列ではありません。", ex);
        }
        if (root == null || !root.isArray()) {
            throw new IllegalArgumentException("担当OP_限定はJSON配列である必要があります。");
        }
        List<String> names = new ArrayList<>();
        LinkedHashSet<String> seen = new LinkedHashSet<>();
        for (JsonNode node : root) {
            if (!node.isTextual()) {
                throw new IllegalArgumentException("担当OP_限定の配列要素は文字列である必要があります。");
            }
            String name = node.textValue();
            if (name == null || name.isBlank()) {
                throw new IllegalArgumentException("担当OP_限定に空のメンバー名は指定できません。");
            }
            if (!seen.add(name)) {
                throw new IllegalArgumentException("担当OP_限定に重複名があります: " + name);
            }
            names.add(name);
        }
        return List.copyOf(names);
    }

    public static String encode(List<String> selectedNames) {
        if (selectedNames == null || selectedNames.isEmpty()) {
            return "";
        }
        LinkedHashSet<String> seen = new LinkedHashSet<>();
        for (String name : selectedNames) {
            if (name == null || name.isBlank()) {
                throw new IllegalArgumentException("担当OP_限定に空のメンバー名は指定できません。");
            }
            if (!seen.add(name)) {
                throw new IllegalArgumentException("担当OP_限定に重複名があります: " + name);
            }
        }
        try {
            return JSON.writeValueAsString(selectedNames);
        } catch (JsonProcessingException ex) {
            throw new IllegalArgumentException("担当OP_限定をJSONへ変換できません。", ex);
        }
    }
}
