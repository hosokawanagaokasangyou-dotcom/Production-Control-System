package jp.co.pm.ai.planning.stage2.source;

import java.util.Objects;
import java.util.concurrent.Callable;
import java.util.concurrent.Executor;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.function.Consumer;

/** 固定ソース確認からFXコールバック完了までを単一flightにする。 */
public final class Stage2SourceGuardCoordinator {
    private final Executor workerExecutor;
    private final Executor uiExecutor;
    private final AtomicBoolean running = new AtomicBoolean();

    public Stage2SourceGuardCoordinator(Executor workerExecutor, Executor uiExecutor) {
        this.workerExecutor = Objects.requireNonNull(workerExecutor);
        this.uiExecutor = Objects.requireNonNull(uiExecutor);
    }

    public boolean isRunning() {
        return running.get();
    }

    public boolean allowsRelatedStart() {
        return !running.get();
    }

    public <T> boolean submit(
            Callable<T> guard,
            Consumer<T> onSucceeded,
            Consumer<Throwable> onFailed) {
        Objects.requireNonNull(guard);
        Objects.requireNonNull(onSucceeded);
        Objects.requireNonNull(onFailed);
        if (!running.compareAndSet(false, true)) {
            return false;
        }
        try {
            workerExecutor.execute(
                    () -> {
                        try {
                            T result = guard.call();
                            dispatchCallback(() -> onSucceeded.accept(result), onFailed);
                        } catch (Throwable failure) {
                            dispatchFailure(failure, onFailed);
                        }
                    });
        } catch (Throwable submissionFailure) {
            running.set(false);
            onFailed.accept(submissionFailure);
            return false;
        }
        return true;
    }

    private void dispatchCallback(Runnable callback, Consumer<Throwable> onFailed) {
        try {
            uiExecutor.execute(
                    () -> {
                        try {
                            callback.run();
                        } catch (Throwable callbackFailure) {
                            onFailed.accept(callbackFailure);
                        } finally {
                            running.set(false);
                        }
                    });
        } catch (Throwable dispatchFailure) {
            running.set(false);
            onFailed.accept(dispatchFailure);
        }
    }

    private void dispatchFailure(Throwable failure, Consumer<Throwable> onFailed) {
        try {
            uiExecutor.execute(
                    () -> {
                        try {
                            onFailed.accept(failure);
                        } finally {
                            running.set(false);
                        }
                    });
        } catch (Throwable dispatchFailure) {
            running.set(false);
            onFailed.accept(dispatchFailure);
        }
    }
}
