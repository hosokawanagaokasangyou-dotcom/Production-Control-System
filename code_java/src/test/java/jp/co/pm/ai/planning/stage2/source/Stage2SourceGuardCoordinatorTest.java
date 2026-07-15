package jp.co.pm.ai.planning.stage2.source;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;

import org.junit.jupiter.api.Test;

class Stage2SourceGuardCoordinatorTest {

    @Test
    void busyRejectsStage1AndMultipleStage2Stage21StartsUntilCallbackReturns()
            throws Exception {
        CountDownLatch guardEntered = new CountDownLatch(1);
        CountDownLatch releaseGuard = new CountDownLatch(1);
        CountDownLatch callbackEntered = new CountDownLatch(1);
        CountDownLatch releaseCallback = new CountDownLatch(1);
        CountDownLatch finished = new CountDownLatch(1);
        Stage2SourceGuardCoordinator coordinator =
                new Stage2SourceGuardCoordinator(
                        command -> {
                            Thread worker = new Thread(command, "guard-test-worker");
                            worker.setDaemon(true);
                            worker.start();
                        },
                        Runnable::run);

        assertTrue(
                coordinator.submit(
                        () -> {
                            guardEntered.countDown();
                            releaseGuard.await(5, TimeUnit.SECONDS);
                            return "ok";
                        },
                        result -> {
                            callbackEntered.countDown();
                            try {
                                releaseCallback.await(5, TimeUnit.SECONDS);
                            } catch (InterruptedException ex) {
                                Thread.currentThread().interrupt();
                            } finally {
                                finished.countDown();
                            }
                        },
                        failure -> finished.countDown()));
        assertTrue(guardEntered.await(5, TimeUnit.SECONDS));
        assertTrue(coordinator.isRunning());
        assertFalse(coordinator.allowsRelatedStart());
        assertFalse(coordinator.submit(() -> "stage2", result -> {}, failure -> {}));
        assertFalse(coordinator.submit(() -> "stage21", result -> {}, failure -> {}));

        releaseGuard.countDown();
        assertTrue(callbackEntered.await(5, TimeUnit.SECONDS));
        assertTrue(coordinator.isRunning());
        releaseCallback.countDown();
        assertTrue(finished.await(5, TimeUnit.SECONDS));
    }

    @Test
    void failureAlwaysReleasesBusy() throws Exception {
        CountDownLatch failed = new CountDownLatch(1);
        Stage2SourceGuardCoordinator coordinator =
                new Stage2SourceGuardCoordinator(
                        command -> {
                            Thread worker = new Thread(command, "guard-test-failure");
                            worker.setDaemon(true);
                            worker.start();
                        },
                        Runnable::run);

        assertTrue(
                coordinator.submit(
                        () -> {
                            throw new IllegalStateException("failure");
                        },
                        result -> {},
                        failure -> failed.countDown()));

        assertTrue(failed.await(5, TimeUnit.SECONDS));
        for (int i = 0; i < 100 && coordinator.isRunning(); i++) {
            Thread.sleep(5);
        }
        assertFalse(coordinator.isRunning());
    }
}
