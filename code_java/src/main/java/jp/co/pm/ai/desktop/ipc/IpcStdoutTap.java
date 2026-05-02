package jp.co.pm.ai.desktop.ipc;

import java.util.function.Consumer;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * Parses NDJSON lines matching ipc-line.schema.json; forwards plain text unchanged.
 */
public final class IpcStdoutTap {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    private IpcStdoutTap() {}

    public static void handleLine(String raw, Consumer<String> humanSummary) {
        if (raw == null || raw.isBlank()) {
            return;
        }
        String t = raw.trim();
        if (!t.startsWith("{")) {
            humanSummary.accept(t);
            return;
        }
        try {
            JsonNode n = MAPPER.readTree(t);
            String type = n.path("type").asText("");
            String msg = switch (type) {
                case "progress" -> n.path("label").asText("...")
                        + " ("
                        + n.path("current").asText("?")
                        + " / "
                        + n.path("total").asText("?")
                        + ")";
                case "log" -> "[" + n.path("level").asText("INFO") + "] " + n.path("message").asText("");
                case "validation_error" -> "[validation_error] " + n.path("message").asText("");
                case "fatal_error" -> "[fatal_error] " + n.path("message").asText("");
                case "done" -> "[done] outputPaths=" + n.path("outputPaths").toString();
                case "ping" -> "[ping] " + n.path("label").asText("");
                default -> t;
            };
            humanSummary.accept(msg);
        } catch (Exception e) {
            humanSummary.accept(t);
        }
    }
}
