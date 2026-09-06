package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicInteger;
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
        // 最後のウィンドウを閉じても FX ランタイムを落とさない（後続テストの runLater が実行されない）
        Platform.runLater(() -> Platform.setImplicitExit(false));
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

    @Test
    void isTabLoadStatus_onlyForBackgroundLoadPhase() {
        assertTrue(
                EnvVarsStartupCheckBusyDialog.isTabLoadStatus(
                        EnvVarsStartupCheckBusyDialog.STATUS_BACKGROUND_LOAD));
        assertTrue(EnvVarsStartupCheckBusyDialog.isTabLoadStatus("起動後読込 (5/6): 原本転記…"));
        assertFalse(
                EnvVarsStartupCheckBusyDialog.isTabLoadStatus(
                        EnvVarsStartupCheckBusyDialog.STATUS_RESTORE_WORKSPACE));
        assertFalse(
                EnvVarsStartupCheckBusyDialog.isTabLoadStatus(
                        EnvVarsStartupCheckBusyDialog.STATUS_MATCH));
        assertFalse(EnvVarsStartupCheckBusyDialog.isTabLoadStatus(null));
    }

    @Test
    void cancelCopy_describesBackgroundContinue() {
        assertEquals("バックグラウンドで続行", EnvVarsStartupCheckBusyDialog.CANCEL_TEXT);
        assertTrue(
                EnvVarsStartupCheckBusyDialog.CANCEL_HINT.contains("バックグラウンド"),
                "ヒントは BG 継続を説明する");
        assertFalse(
                EnvVarsStartupCheckBusyDialog.CANCEL_HINT.contains("中断"),
                "ヒントは読込中断を約束しない");
    }

    @Test
    void cancel_isEnabledOnlyDuringTabLoad_andRunsHandler() throws Exception {
        AtomicReference<EnvVarsStartupCheckBusyDialog> ref = new AtomicReference<>();
        AtomicInteger cancelled = new AtomicInteger();
        CountDownLatch shown = new CountDownLatch(1);
        Platform.runLater(
                () -> {
                    ref.set(
                            EnvVarsStartupCheckBusyDialog.show(
                                    null,
                                    EnvVarsStartupCheckBusyDialog.STATUS_RESTORE_WORKSPACE,
                                    cancelled::incrementAndGet));
                    shown.countDown();
                });
        assertTrue(shown.await(10, TimeUnit.SECONDS));
        EnvVarsStartupCheckBusyDialog dialog = ref.get();
        assertFalse(dialog.isCancelEnabled(), "必須チェック中はキャンセル不可");

        CountDownLatch done = new CountDownLatch(1);
        Platform.runLater(
                () -> {
                    dialog.setCancelEnabled(true);
                    dialog.fireCancelForTest();
                    dialog.close();
                    done.countDown();
                });
        assertTrue(done.await(10, TimeUnit.SECONDS));
        assertEquals(1, cancelled.get(), "タブ読込段階ではキャンセル処理が走る");
        assertFalse(dialog.isShowing());
    }
}
