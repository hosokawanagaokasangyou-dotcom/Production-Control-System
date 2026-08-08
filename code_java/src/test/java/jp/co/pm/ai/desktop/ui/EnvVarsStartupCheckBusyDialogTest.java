package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledOnOs;
import org.junit.jupiter.api.condition.OS;

import javafx.application.Platform;

@EnabledOnOs(OS.WINDOWS)
class EnvVarsStartupCheckBusyDialogTest {

    @BeforeAll
    static void initFx() throws Exception {
        CountDownLatch ready = new CountDownLatch(1);
        try {
            Platform.startup(ready::countDown);
        } catch (IllegalStateException alreadyStarted) {
            ready.countDown();
        }
        assertTrue(ready.await(10, TimeUnit.SECONDS));
    }

    @Test
    void show_updatesStatus_andCloses() throws Exception {
        AtomicReference<EnvVarsStartupCheckBusyDialog> ref = new AtomicReference<>();
        CountDownLatch shown = new CountDownLatch(1);
        Platform.runLater(
                () -> {
                    EnvVarsStartupCheckBusyDialog dialog =
                            EnvVarsStartupCheckBusyDialog.show(
                                    null, EnvVarsStartupCheckBusyDialog.STATUS_STABILIZE);
                    ref.set(dialog);
                    dialog.setStatus(EnvVarsStartupCheckBusyDialog.STATUS_MATCH);
                    dialog.setStep(EnvVarsStartupCheckBusyDialog.STEP_ENV_MATCH);
                    dialog.setHeader(EnvVarsStartupCheckBusyDialog.HEADER_BACKGROUND_LOAD);
                    shown.countDown();
                });
        assertTrue(shown.await(10, TimeUnit.SECONDS));

        EnvVarsStartupCheckBusyDialog dialog = ref.get();
        assertTrue(dialog.isShowing());

        CountDownLatch closed = new CountDownLatch(1);
        Platform.runLater(
                () -> {
                    dialog.close();
                    closed.countDown();
                });
        assertTrue(closed.await(10, TimeUnit.SECONDS));
        assertFalse(dialog.isShowing());
        assertEquals("起動時チェック", EnvVarsStartupCheckBusyDialog.TITLE);
    }
}
