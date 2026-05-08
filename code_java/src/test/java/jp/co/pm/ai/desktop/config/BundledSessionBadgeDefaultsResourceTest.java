package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertNotNull;

import org.junit.jupiter.api.Test;

class BundledSessionBadgeDefaultsResourceTest {

    @Test
    void bundledSessionBadgeDefaults_jsonOnClasspath() {
        assertNotNull(
                DesktopSessionStateStore.class.getResource(
                        "/jp/co/pm/ai/desktop/config/bundled_session_badge_defaults.json"));
    }
}
