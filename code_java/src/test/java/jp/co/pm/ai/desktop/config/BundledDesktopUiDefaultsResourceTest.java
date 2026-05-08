package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertNotNull;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;

class BundledDesktopUiDefaultsResourceTest {

    @Test
    void bundledSessionUiDefaults_jsonOnClasspath() {
        assertNotNull(
                DesktopSessionStateStore.class.getResource(
                        "/jp/co/pm/ai/desktop/config/bundled_session_ui_defaults.json"));
    }

    @Test
    void bundledTableColumnOrder_jsonOnClasspath() {
        assertNotNull(
                TableColumnOrderPersistence.class.getResource(
                        "/jp/co/pm/ai/desktop/config/bundled_table_column_order.json"));
    }
}
