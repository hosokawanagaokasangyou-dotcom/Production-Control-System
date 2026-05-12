package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertFalse;

import org.junit.jupiter.api.Test;

class UiRefEnvDefaultsTest {

    @Test
    void load_contains_entries_from_ui_ref_snapshot() {
        assertFalse(UiRefEnvDefaults.loadOrEmpty().isEmpty(), "run GenerateUiRefEnvDefaultsJson if empty");
    }
}
