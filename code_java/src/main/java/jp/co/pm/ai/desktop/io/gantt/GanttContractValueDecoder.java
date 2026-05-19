package jp.co.pm.ai.desktop.io.gantt;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.List;

import com.fasterxml.jackson.databind.JsonNode;

/**
 * {@code gantt_render_contract} の {@code encode_value} 互換の JSON ノードを
 * 実行時型へ戻す（date / datetime / time / tuple のみ対応）。
 */
public final class GanttContractValueDecoder {

    private GanttContractValueDecoder() {}

    public static Object decodeValue(JsonNode n) {
        if (n == null || n.isNull()) {
            return null;
        }
        if (n.isObject()) {
            JsonNode t = n.get("__t");
            if (t != null && t.isTextual()) {
                String kind = t.asText();
                if ("date".equals(kind)) {
                    return LocalDate.parse(n.get("v").asText());
                }
                if ("datetime".equals(kind)) {
                    String s = n.get("v").asText();
                    if (s.length() >= 19) {
                        return LocalDateTime.parse(s.substring(0, 19));
                    }
                    return LocalDateTime.parse(s);
                }
                if ("time".equals(kind)) {
                    String s = n.get("v").asText();
                    if (s.length() >= 5) {
                        return LocalTime.parse(s.length() > 8 ? s.substring(0, 8) : s);
                    }
                }
                if ("tuple".equals(kind)) {
                    JsonNode items = n.get("items");
                    List<Object> out = new ArrayList<>();
                    if (items != null && items.isArray()) {
                        for (JsonNode x : items) {
                            out.add(decodeValue(x));
                        }
                    }
                    return out;
                }
            }
            return null;
        }
        if (n.isArray()) {
            List<Object> out = new ArrayList<>();
            for (JsonNode x : n) {
                out.add(decodeValue(x));
            }
            return out;
        }
        if (n.isTextual()) {
            return n.asText();
        }
        if (n.isBoolean()) {
            return n.asBoolean();
        }
        if (n.isInt()) {
            return n.asInt();
        }
        if (n.isLong()) {
            return n.asLong();
        }
        if (n.isDouble()) {
            return n.asDouble();
        }
        return null;
    }

    public static LocalDateTime toLocalDateTime(Object o) {
        if (o instanceof LocalDateTime ldt) {
            return ldt;
        }
        return null;
    }

    public static LocalDate toLocalDate(Object o) {
        if (o instanceof LocalDate ld) {
            return ld;
        }
        return null;
    }

    public static LocalTime toLocalTime(Object o) {
        if (o instanceof LocalTime lt) {
            return lt;
        }
        return null;
    }
}
