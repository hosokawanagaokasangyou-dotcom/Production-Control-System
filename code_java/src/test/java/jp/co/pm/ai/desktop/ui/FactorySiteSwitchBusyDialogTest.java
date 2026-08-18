package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;

import javafx.application.Platform;
import javafx.stage.StageStyle;

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
        AtomicReference<Throwable> fxError = new AtomicReference<>();
        CountDownLatch shown = new CountDownLatch(1);
        Platform.runLater(
                () -> {
                    try {
                        FactorySiteSwitchBusyDialog dialog =
                                FactorySiteSwitchBusyDialog.show(
                                        null,
                                        "湖南工場 → 国分工場",
                                        FactorySiteSwitchBusyDialog.STATUS_SAVING);
                        ref.set(dialog);
                        dialog.setStatus(FactorySiteSwitchBusyDialog.STATUS_ENV);
                    } catch (Throwable t) {
                        fxError.set(t);
                    } finally {
                        shown.countDown();
                    }
                });
        assertTrue(shown.await(10, TimeUnit.SECONDS), "FX スレッドで show が完了しない");
        assertEquals(null, fxError.get());

        FactorySiteSwitchBusyDialog dialog = ref.get();
        assertNotNull(dialog);
        assertTrue(dialog.isShowing());
        assertEquals(StageStyle.UNDECORATED, dialog.stageStyle());
        assertEquals(
                FactorySiteSwitchBusyDialog.STATUS_BACKGROUND_LOAD,
                FactorySiteSwitchBusySupport.resolveTabLoadStatus(""));

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
