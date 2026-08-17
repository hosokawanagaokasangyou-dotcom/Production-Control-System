package jp.co.pm.ai.desktop.ui;

import java.util.concurrent.ThreadLocalRandom;
import java.util.function.IntUnaryOperator;

/**
 * 終了ゲート用の7桁確認コード（1000000〜9999999）。
 */
public final class SevenDigitChallenge {

    private SevenDigitChallenge() {}

    public static String generate() {
        return generate(ThreadLocalRandom.current()::nextInt);
    }

    public static String generate(IntUnaryOperator randomBound) {
        IntUnaryOperator rnd = randomBound != null ? randomBound : bound -> 0;
        int n = 1_000_000 + rnd.applyAsInt(9_000_000);
        return Integer.toString(n);
    }

    public static boolean matches(String expected, String input) {
        return expected != null && input != null && expected.equals(input.strip());
    }
}
