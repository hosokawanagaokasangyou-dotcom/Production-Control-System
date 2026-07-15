package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;

import javafx.application.Platform;
import javafx.scene.control.ButtonBar;
import javafx.scene.control.Dialog;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

class LimitedOperatorCellEditorTest {

    @BeforeAll
    static void initJavaFx() {
        try {
            Platform.startup(() -> {});
        } catch (IllegalStateException ignored) {
            // already started
        }
    }

    @Test
    void loadingDialogHasCancelCloseButton() throws Exception {
        CountDownLatch completed = new CountDownLatch(1);
        AtomicReference<Dialog<Void>> dialogRef = new AtomicReference<>();
        Platform.runLater(
                () -> {
                    dialogRef.set(LimitedOperatorCellEditor.createBusyDialog(null));
                    completed.countDown();
                });

        assertTrue(completed.await(5, TimeUnit.SECONDS));
        Dialog<Void> dialog = dialogRef.get();
        assertEquals(1, dialog.getDialogPane().getButtonTypes().size());
        assertEquals(
                ButtonBar.ButtonData.CANCEL_CLOSE,
                dialog.getDialogPane().getButtonTypes().getFirst().getButtonData());
    }
}
