package jp.co.pm.ai.desktop.ui;

import java.util.concurrent.ThreadLocalRandom;
import java.util.function.LongUnaryOperator;

/**
 * 終了ゲート用の4桁確認コード（1000〜9999）。
 */
public final class SevenDigitChallenge {

    public static final int DIGIT_COUNT = 4;

    private static final long MIN = 1000L;

    private static final long SPAN = 9000L;

    private SevenDigitChallenge() {}

    public static String generate() {
        return generate(bound -> ThreadLocalRandom.current().nextLong(bound));
    }

    public static String generate(LongUnaryOperator randomBound) {
        LongUnaryOperator rnd = randomBound != null ? randomBound : bound -> 0L;
        long n = MIN + rnd.applyAsLong(SPAN);
        return Long.toString(n);
    }

    public static boolean matches(String expected, String input) {
        return expected != null && input != null && expected.equals(input.strip());
    }
}
