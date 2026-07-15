package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.Executor;
import java.util.concurrent.atomic.AtomicReference;

import org.junit.jupiter.api.Test;

class LimitedOperatorLoadCoordinatorTest {

    @Test
    void submitPreventsConcurrentLoadsAndReleasesBusyBeforeSuccessCallback() {
        List<Runnable> workerQueue = new ArrayList<>();
        Executor queuedWorker = workerQueue::add;
        AtomicReference<String> result = new AtomicReference<>();
        LimitedOperatorLoadCoordinator coordinator = new LimitedOperatorLoadCoordinator();

        assertTrue(
                coordinator.submit(
                        () -> "候補",
                        queuedWorker,
                        Runnable::run,
                        value -> {
                            assertFalse(coordinator.isBusy());
                            result.set(value);
                        },
                        failure -> {}));
        assertTrue(coordinator.isBusy());
        assertFalse(
                coordinator.submit(
                        () -> "二重起動",
                        queuedWorker,
                        Runnable::run,
                        result::set,
                        failure -> {}));

        workerQueue.removeFirst().run();

        assertEquals("候補", result.get());
        assertFalse(coordinator.isBusy());
    }

    @Test
    void failedLoadReleasesBusyAndCallsOnlyFailure() {
        AtomicReference<Throwable> failure = new AtomicReference<>();
        AtomicReference<String> result = new AtomicReference<>();
        LimitedOperatorLoadCoordinator coordinator = new LimitedOperatorLoadCoordinator();

        assertTrue(
                coordinator.submit(
                        () -> {
                            throw new IllegalStateException("読込失敗");
                        },
                        Runnable::run,
                        Runnable::run,
                        result::set,
                        failure::set));

        assertEquals(null, result.get());
        assertEquals("読込失敗", failure.get().getMessage());
        assertFalse(coordinator.isBusy());
    }
}
