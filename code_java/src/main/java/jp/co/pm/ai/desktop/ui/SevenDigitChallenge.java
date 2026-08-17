package jp.co.pm.ai.desktop.ui;

import java.util.concurrent.ThreadLocalRandom;
import java.util.function.LongUnaryOperator;

/**
 * 終了ゲート用の12桁確認コード（100000000000〜999999999999）。
 */
public final class SevenDigitChallenge {

    public static final int DIGIT_COUNT = 12;

    private static final long MIN = 100_000_000_000L;

    private static final long SPAN = 900_000_000_000L;

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
