package jp.co.pm.ai.desktop.ui;

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
class StageRunBusyDialogTest {

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
    void show_updatesPhaseDetail_andCloses() throws Exception {
        AtomicReference<StageRunBusyDialog> ref = new AtomicReference<>();
        CountDownLatch shown = new CountDownLatch(1);
        Platform.runLater(
                () -> {
                    StageRunBusyDialog dialog =
                            StageRunBusyDialog.show(
                                    null,
                                    "段階2 実行中",
                                    "段階2（配台計画）を実行しています",
                                    "準備中…",
                                    () -> {});
                    ref.set(dialog);
                    dialog.setPhase("Python 実行中…");
                    dialog.setDetail("段階2: 計画シミュレーションを開始");
                    shown.countDown();
                });
        assertTrue(shown.await(10, TimeUnit.SECONDS));

        StageRunBusyDialog dialog = ref.get();
        assertTrue(dialog.isShowing());

        CountDownLatch closed = new CountDownLatch(1);
        Platform.runLater(
                () -> {
                    dialog.close();
                    closed.countDown();
                });
        assertTrue(closed.await(10, TimeUnit.SECONDS));
        assertFalse(dialog.isShowing());
    }
}
