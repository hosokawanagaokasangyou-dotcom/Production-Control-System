package jp.co.pm.ai.desktop.dispatch.rules.migration;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

class DispatchRuleMigrationServiceTest {

    private static final ObjectMapper JSON = new ObjectMapper();

    @Test
    void v0PriorityBecomesApplyOrder() throws Exception {
        ObjectNode raw = JSON.createObjectNode();
        raw.put("version", 0);
        var rules = raw.putArray("rules");
        var rule = rules.addObject();
        rule.put("id", "L13");
        rule.put("priority", 40);
        ObjectNode migrated = DispatchRuleMigrationService.migrate(raw);
        assertEquals(1, migrated.get("schemaVersion").asInt());
        assertEquals(40, migrated.get("rules").get(0).get("applyOrder").asInt());
        assertTrue(migrated.get("rules").get(0).has("enabled"));
    }
}
