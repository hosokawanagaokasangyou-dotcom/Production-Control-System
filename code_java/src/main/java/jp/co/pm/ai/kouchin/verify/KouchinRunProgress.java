package jp.co.pm.ai.kouchin.verify;

import java.util.concurrent.atomic.AtomicBoolean;
import java.util.function.Consumer;

/**
 * 後加工工賃の実行中に、いま何をしているかを画面へ渡す。
 * リスナーは UI スレッドへ載せること。
 */
public final class KouchinRunProgress {

    private static volatile Consumer<String> listener = msg -> {};
    private static volatile AtomicBoolean cancel = new AtomicBoolean(false);

    private KouchinRunProgress() {}

    public static void bind(Consumer<String> onMessage, AtomicBoolean cancelFlag) {
        listener = onMessage == null ? msg -> {} : onMessage;
        cancel = cancelFlag == null ? new AtomicBoolean(false) : cancelFlag;
    }

    public static void clear() {
        listener = msg -> {};
        cancel = new AtomicBoolean(false);
    }

    public static void report(String message) {
        if (message == null || message.isBlank()) {
            return;
        }
        listener.accept(message);
    }

    public static boolean cancelled() {
        AtomicBoolean flag = cancel;
        return flag != null && flag.get();
    }
}
