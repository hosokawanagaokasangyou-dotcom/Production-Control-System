package jp.co.pm.ai.desktop.io.gantt;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import com.fasterxml.jackson.databind.JsonNode;

import jp.co.pm.ai.desktop.debug.AgentDebugLog;

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
                    try {
                        if (s.length() >= 19) {
                            return LocalDateTime.parse(s.substring(0, 19));
                        }
                        return LocalDateTime.parse(s);
                    } catch (RuntimeException ex) {
                        // #region agent log
                        Map<String, Object> d = new LinkedHashMap<>();
                        d.put(
                                "raw",
                                s == null
                                        ? ""
                                        : (s.length() > 120 ? s.substring(0, 120) : s));
                        d.put("err", ex.getClass().getSimpleName() + ":" + ex.getMessage());
                        AgentDebugLog.appendStructured(
                                Map.of(),
                                "1ecccd",
                                "B",
                                "GanttContractValueDecoder.decodeValue:datetime",
                                "LocalDateTime.parse failed",
                                d);
                        // #endregion
                        throw ex;
                    }
                }
                if ("time".equals(kind)) {
                    String s = n.get("v").asText();
                    if (s.length() >= 5) {
                        String frag = s.length() > 8 ? s.substring(0, 8) : s;
                        try {
                            return LocalTime.parse(frag);
                        } catch (RuntimeException ex) {
                            // #region agent log
                            Map<String, Object> d = new LinkedHashMap<>();
                            d.put("raw", s.length() > 120 ? s.substring(0, 120) : s);
                            d.put("frag", frag);
                            d.put("err", ex.getClass().getSimpleName() + ":" + ex.getMessage());
                            AgentDebugLog.appendStructured(
                                    Map.of(),
                                    "1ecccd",
                                    "C",
                                    "GanttContractValueDecoder.decodeValue:time",
                                    "LocalTime.parse failed",
                                    d);
                            // #endregion
                            throw ex;
                        }
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
}
