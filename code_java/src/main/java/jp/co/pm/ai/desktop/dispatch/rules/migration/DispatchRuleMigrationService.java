package jp.co.pm.ai.desktop.dispatch.rules.migration;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

/** Schema migrations aligned with Python {@code migrations.py}. */
public final class DispatchRuleMigrationService {

    public static final int CURRENT_SCHEMA_VERSION = 1;
    public static final int SUPPORTED_SCHEMA_MAX = 1;

    private static final ObjectMapper JSON = new ObjectMapper();

    private DispatchRuleMigrationService() {}

    public static ObjectNode migrate(ObjectNode raw) {
        ObjectNode doc = raw.deepCopy();
        int ver = detectVersion(doc);
        if (ver > SUPPORTED_SCHEMA_MAX) {
            throw new IllegalArgumentException(
                    "schemaVersion " + ver + " exceeds supported max " + SUPPORTED_SCHEMA_MAX);
        }
        while (ver < CURRENT_SCHEMA_VERSION) {
            if (ver == 0) {
                doc = applyV0ToV1(doc);
            } else {
                throw new IllegalArgumentException("No migration from schemaVersion " + ver);
            }
            ver = detectVersion(doc);
        }
        doc.put("schemaVersion", CURRENT_SCHEMA_VERSION);
        return doc;
    }

    private static int detectVersion(ObjectNode doc) {
        if (doc.has("schemaVersion")) {
            return doc.get("schemaVersion").asInt(0);
        }
        if (doc.has("version")) {
            return doc.get("version").asInt(0);
        }
        return 0;
    }

    private static ObjectNode applyV0ToV1(ObjectNode doc) {
        doc.put("schemaVersion", 1);
        doc.remove("version");
        JsonNode rules = doc.get("rules");
        if (!(rules instanceof ArrayNode arr)) {
            doc.set("rules", JSON.createArrayNode());
            return doc;
        }
        for (JsonNode n : arr) {
            if (n instanceof ObjectNode rule) {
                if (!rule.has("applyOrder") && rule.has("priority")) {
                    rule.put("applyOrder", rule.get("priority").asInt(100));
                    rule.remove("priority");
                }
                if (!rule.has("enabled")) {
                    rule.put("enabled", true);
                }
                if (!rule.has("executionMode")) {
                    rule.put("executionMode", "auto");
                }
                if (!rule.has("legacyFallback")) {
                    rule.put("legacyFallback", true);
                }
                if (!rule.has("graph")) {
                    rule.set("graph", JSON.createObjectNode());
                }
            }
        }
        return doc;
    }
}
