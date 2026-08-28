package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;

import javafx.application.Platform;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

class MainRunTabControllerLogRestoreTest {

    @BeforeAll
    static void initJavaFx() {
        try {
            Platform.startup(() -> {});
        } catch (IllegalStateException ignored) {
            // already started
        }
    }

    @Test
    void preservesUpgradeLogsAcrossShellUiRestore() throws Exception {
        CountDownLatch completed = new CountDownLatch(1);
        AtomicReference<AssertionError> failure = new AtomicReference<>();
        Platform.runLater(
                () -> {
                    try {
                        MainRunTabController controller = new MainRunTabController();
                        controller.appendLog("[env] 環境変数を ui_ref 既定に初期化しました。", false);
                        List<String> preserved =
                                controller.snapshotLogLinesForShellUiRestore();

                        controller.clearMainRunTabLog();
                        controller.restoreLogLinesAfterShellUiRestore(preserved);

                        assertEquals(
                                List.of("[env] 環境変数を ui_ref 既定に初期化しました。"),
                                controller.snapshotLogLinesForShellUiRestore());
                    } catch (AssertionError error) {
                        failure.set(error);
                    } finally {
                        completed.countDown();
                    }
                });

        assertTrue(completed.await(5, TimeUnit.SECONDS));
        if (failure.get() != null) {
            throw failure.get();
        }
    }
}
