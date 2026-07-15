package jp.co.pm.ai.desktop.ui;

import java.util.Objects;
import java.util.concurrent.Callable;
import java.util.concurrent.Executor;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.function.Consumer;

/** 資格候補の非同期読込について、多重起動防止と完了スレッド切替を担う。 */
public final class LimitedOperatorLoadCoordinator {

    private final AtomicBoolean busy = new AtomicBoolean(false);

    public boolean isBusy() {
        return busy.get();
    }

    public <T> boolean submit(
            Callable<T> loader,
            Executor workerExecutor,
            Executor completionExecutor,
            Consumer<T> onSuccess,
            Consumer<Throwable> onFailure) {
        Objects.requireNonNull(loader, "loader");
        Objects.requireNonNull(workerExecutor, "workerExecutor");
        Objects.requireNonNull(completionExecutor, "completionExecutor");
        Objects.requireNonNull(onSuccess, "onSuccess");
        Objects.requireNonNull(onFailure, "onFailure");
        if (!busy.compareAndSet(false, true)) {
            return false;
        }
        try {
            workerExecutor.execute(
                    () -> {
                        try {
                            T result = loader.call();
                            dispatchCompletion(
                                    completionExecutor,
                                    () -> onSuccess.accept(result),
                                    onFailure);
                        } catch (Throwable failure) {
                            dispatchCompletion(
                                    completionExecutor,
                                    () -> onFailure.accept(failure),
                                    onFailure);
                        }
                    });
        } catch (Throwable failure) {
            busy.set(false);
            onFailure.accept(failure);
        }
        return true;
    }

    private void dispatchCompletion(
            Executor completionExecutor,
            Runnable completion,
            Consumer<Throwable> onFailure) {
        try {
            completionExecutor.execute(
                    () -> {
                        busy.set(false);
                        completion.run();
                    });
        } catch (Throwable failure) {
            busy.set(false);
            onFailure.accept(failure);
        }
    }
}
