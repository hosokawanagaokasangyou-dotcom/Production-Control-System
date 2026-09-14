package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;

import javafx.application.Platform;
import javafx.scene.control.Button;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledOnOs;
import org.junit.jupiter.api.condition.OS;

@EnabledOnOs(OS.WINDOWS)
class ButtonAttentionGlowTest {

    @BeforeAll
    static void initJavaFx() {
        try {
            Platform.startup(() -> {});
        } catch (IllegalStateException ignored) {
            // already started
        }
    }

    @Test
    void unsavedSaveGlowTurnsOnAndOffWithDirtyFlag() throws Exception {
        AtomicReference<Button> buttonRef = new AtomicReference<>();
        AtomicReference<ButtonAttentionGlow> glowRef = new AtomicReference<>();
        AtomicReference<Throwable> fxError = new AtomicReference<>();
        CountDownLatch done = new CountDownLatch(1);
        Platform.runLater(
                () -> {
                    try {
                        Button save = new Button("保存");
                        ButtonAttentionGlow glow = ButtonAttentionGlow.forUnsavedSave(save);
                        glow.apply(true);
                        buttonRef.set(save);
                        glowRef.set(glow);
                    } catch (Throwable t) {
                        fxError.set(t);
                    } finally {
                        done.countDown();
                    }
                });
        assertTrue(done.await(15, TimeUnit.SECONDS), "FX 初期化が完了しない");
        if (fxError.get() != null) {
            throw new AssertionError(fxError.get());
        }
        Button save = buttonRef.get();
        assertTrue(
                save.getStyleClass().contains(ButtonAttentionGlow.UNSAVED_SAVE_STYLE_CLASS),
                "未保存時は保存ボタン用の強調クラスが付く");

        CountDownLatch off = new CountDownLatch(1);
        Platform.runLater(
                () -> {
                    glowRef.get().apply(false);
                    off.countDown();
                });
        assertTrue(off.await(15, TimeUnit.SECONDS));
        assertFalse(
                save.getStyleClass().contains(ButtonAttentionGlow.UNSAVED_SAVE_STYLE_CLASS),
                "保存後は強調クラスが外れる");
    }
}
