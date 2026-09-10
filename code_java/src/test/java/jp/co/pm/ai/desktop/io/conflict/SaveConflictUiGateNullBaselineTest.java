package jp.co.pm.ai.desktop.io.conflict;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.concurrent.atomic.AtomicReference;

import org.junit.jupiter.api.Test;

class SaveConflictUiGateNullBaselineTest {

    @Test
    void nullBaseline_abortsSave() {
        AtomicReference<String> err = new AtomicReference<>();
        boolean ok =
                SaveConflictUiGate.allowSave(
                        null,
                        "テスト",
                        null,
                        (a, b, c) -> "",
                        () -> {},
                        err::set);
        assertFalse(ok);
        assertTrue(err.get() != null && err.get().contains("指紋"), err.get());
    }
}
