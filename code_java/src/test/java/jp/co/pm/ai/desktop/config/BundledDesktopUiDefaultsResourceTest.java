package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.InputStream;

import org.junit.jupiter.api.Test;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;

class BundledDesktopUiDefaultsResourceTest {

    @Test
    void bundledSessionUiDefaults_jsonOnClasspath() {
        assertNotNull(
                DesktopSessionStateStore.class.getResource(
                        "/jp/co/pm/ai/desktop/config/bundled_session_ui_defaults.json"));
    }

    @Test
    void bundledSessionUiDefaults_containsMainShellTabOrder() throws Exception {
        try (InputStream in =
                DesktopSessionStateStore.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/config/bundled_session_ui_defaults.json")) {
            assertNotNull(in);
            JsonNode n = new ObjectMapper().readTree(in);
            assertTrue(n.path("mainShellTabOrder").isArray());
            assertEquals(18, n.path("mainShellTabOrder").size());
            assertTrue(n.path("mainShellTabTitleAliases").isObject());
        }
    }

    @Test
    void bundledTableColumnOrder_jsonOnClasspath() {
        assertNotNull(
                TableColumnOrderPersistence.class.getResource(
                        "/jp/co/pm/ai/desktop/config/bundled_table_column_order.json"));
    }
}
