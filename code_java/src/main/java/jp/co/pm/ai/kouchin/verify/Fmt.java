package jp.co.pm.ai.kouchin.verify;

import java.util.Locale;

/**
 * レポート文言用の金額フォーマット（Python の {@code f'{v:,.0f}'} 相当）。
 */
public final class Fmt {

    private Fmt() {
    }

    /** {@code 1,234} 形式（小数切り捨てなし・四捨五入）。 */
    public static String n0(double value) {
        return String.format(Locale.US, "%,.0f", zeroClean(value));
    }

    /** {@code +1,234} / {@code -1,234} 形式。 */
    public static String s0(double value) {
        return String.format(Locale.US, "%+,.0f", zeroClean(value));
    }

    /** {@code 1,234.56} 形式。 */
    public static String n2(double value) {
        return String.format(Locale.US, "%,.2f", zeroClean(value));
    }

    /** 小数第2位で丸める（合算で生じる {@code -0.0000001} を消す）。 */
    public static double round2(double value) {
        double r = Math.round(value * 100.0) / 100.0;
        return r == 0.0 ? 0.0 : r;
    }

    /** 整数丸め（Python の {@code round()} 相当・銀行丸めではなく四捨五入）。 */
    public static long round0(double value) {
        return Math.round(zeroClean(value));
    }

    private static double zeroClean(double value) {
        return value == 0.0 ? 0.0 : value;
    }
}
