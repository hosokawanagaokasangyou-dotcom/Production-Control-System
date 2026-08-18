package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class IdentityCheckHistoryTabControllerTest {

    @Test
    void resultLabel_mapsKnownKeys() {
        assertEquals("同一", IdentityCheckHistoryTabController.resultLabel("ok"));
        assertEquals("差異", IdentityCheckHistoryTabController.resultLabel("mismatch"));
        assertEquals("", IdentityCheckHistoryTabController.resultLabel(null));
        assertEquals("other", IdentityCheckHistoryTabController.resultLabel("other"));
    }

    @Test
    void formatTs_formatsOffsetDateTime() {
        assertEquals(
                "2026-08-18 20:45:12",
                IdentityCheckHistoryTabController.formatTs("2026-08-18T20:45:12+09:00"));
        assertEquals("", IdentityCheckHistoryTabController.formatTs(""));
        assertEquals("not-a-date", IdentityCheckHistoryTabController.formatTs("not-a-date"));
    }
}
