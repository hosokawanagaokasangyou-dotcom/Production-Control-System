package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;

import javafx.application.Platform;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledOnOs;
import org.junit.jupiter.api.condition.OS;

@EnabledOnOs(OS.WINDOWS)
class FactorySiteSwitchBusyDialogTest {

    @BeforeAll
    static void initFx() {
        try {
            Platform.startup(() -> {});
        } catch (IllegalStateException alreadyStarted) {
            // already started
        }
    }

    @Test
    void show_updatesStatus_andCloses() throws Exception {
        AtomicReference<FactorySiteSwitchBusyDialog> ref = new AtomicReference<>();
        CountDownLatch shown = new CountDownLatch(1);
        Platform.runLater(
                () -> {
                    FactorySiteSwitchBusyDialog dialog =
                            FactorySiteSwitchBusyDialog.show(
                                    null,
                                    "湖南工場 → 国分工場",
                                    FactorySiteSwitchBusyDialog.STATUS_SAVING);
                    ref.set(dialog);
                    dialog.setStatus(FactorySiteSwitchBusyDialog.STATUS_ENV);
                    shown.countDown();
                });
        assertTrue(shown.await(5, TimeUnit.SECONDS));

        FactorySiteSwitchBusyDialog dialog = ref.get();
        assertTrue(dialog.isShowing());

        CountDownLatch closed = new CountDownLatch(1);
        Platform.runLater(
                () -> {
                    dialog.close();
                    closed.countDown();
                });
        assertTrue(closed.await(5, TimeUnit.SECONDS));
        assertFalse(dialog.isShowing());
        assertEquals("工場切替", FactorySiteSwitchBusyDialog.TITLE);
    }
}
