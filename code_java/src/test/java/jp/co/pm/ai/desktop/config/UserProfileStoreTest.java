package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;

import org.junit.jupiter.api.Test;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

class UserProfileStoreTest {

    private static final ObjectMapper JSON = new ObjectMapper();

    @Test
    void loadProfileFromJson_minimalV1_loads() throws Exception {
        ObjectNode root = JSON.createObjectNode();
        root.put("schemaVersion", 1);
        root.put("savedAt", "2020-01-01T12:00:00");
        root.put("displayName", "unit");
        root.set("session", JSON.createObjectNode());
        root.set("tableColumnOrder", JSON.createObjectNode());

        UserProfileStore.LoadedProfile loaded = UserProfileStore.loadProfileFromJson(root);
        assertEquals("unit", loaded.displayName());
        assertDoesNotThrow(() -> loaded.session().planInputPath());
    }

    @Test
    void loadProfileFromJson_schemaVersionMissing_migratesFromZero() throws Exception {
        ObjectNode root = JSON.createObjectNode();
        root.set("session", JSON.createObjectNode());
        root.set("tableColumnOrder", JSON.createObjectNode());

        UserProfileStore.LoadedProfile loaded = UserProfileStore.loadProfileFromJson(root);
        assertTrue(root.path("schemaVersion").asInt(0) >= 1);
        assertDoesNotThrow(() -> loaded.session().uiTheme());
    }

    @Test
    void loadProfileFromJson_unknownSessionKeys_ignored() throws Exception {
        ObjectNode session = JSON.createObjectNode();
        session.put("futureUnknownFlag_xyz", true);
        session.put("planInputPath", "/tmp/example.xls");

        ObjectNode root = JSON.createObjectNode();
        root.put("schemaVersion", 1);
        root.set("session", session);
        root.set("tableColumnOrder", JSON.createObjectNode());

        UserProfileStore.LoadedProfile loaded = UserProfileStore.loadProfileFromJson(root);
        assertEquals("/tmp/example.xls", loaded.session().planInputPath());
    }

    @Test
    void loadProfileFromJson_futureEnvelopeSchema_rejected() {
        ObjectNode root = JSON.createObjectNode();
        root.put("schemaVersion", UserProfileStore.SUPPORTED_PROFILE_SCHEMA_MAX + 1);
        root.set("session", JSON.createObjectNode());
        root.set("tableColumnOrder", JSON.createObjectNode());

        IOException ex =
                assertThrows(IOException.class, () -> UserProfileStore.loadProfileFromJson(root));
        assertTrue(ex.getMessage().contains("schemaVersion"));
    }

    @Test
    void loadProfileFromJson_sessionNotObject_throws() {
        ObjectNode root = JSON.createObjectNode();
        root.put("schemaVersion", 1);
        root.put("session", "invalid");

        assertThrows(IOException.class, () -> UserProfileStore.loadProfileFromJson(root));
    }
}
