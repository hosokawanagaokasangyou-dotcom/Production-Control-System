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
class PostProcessingMasterLoadBusyDialogTest {

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
        AtomicReference<PostProcessingMasterLoadBusyDialog> ref = new AtomicReference<>();
        AtomicReference<Throwable> fxError = new AtomicReference<>();
        CountDownLatch shown = new CountDownLatch(1);
        Platform.runLater(
                () -> {
                    try {
                        PostProcessingMasterLoadBusyDialog dialog =
                                PostProcessingMasterLoadBusyDialog.show(
                                        null, PostProcessingMasterLoadBusyDialog.STATUS_LOADING);
                        ref.set(dialog);
                        dialog.setStatus("見出しを読み込んでいます…");
                    } catch (Throwable t) {
                        fxError.set(t);
                    } finally {
                        shown.countDown();
                    }
                });
        assertTrue(shown.await(10, TimeUnit.SECONDS), "FX スレッドで show が完了しない");
        assertEquals(null, fxError.get());

        PostProcessingMasterLoadBusyDialog dialog = ref.get();
        assertNotNull(dialog);
        assertTrue(dialog.isShowing());
        assertEquals(StageStyle.UNDECORATED, dialog.stageStyle());
        assertTrue(dialog.scene().getStylesheets().isEmpty());
        assertTrue(
                dialog.bodyTextStyle().contains(PostProcessingMasterLoadBusyDialog.BODY_TEXT_FILL),
                "明るい背景向けに本文は暗い文字色にする");

        CountDownLatch closed = new CountDownLatch(1);
        Platform.runLater(
                () -> {
                    dialog.close();
                    closed.countDown();
                });
        assertTrue(closed.await(5, TimeUnit.SECONDS));
        assertFalse(dialog.isShowing());
        assertEquals("後加工商品マスタ", PostProcessingMasterLoadBusyDialog.TITLE);
    }
}
