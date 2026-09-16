package jp.co.pm.ai.kouchin.verify;

import java.math.BigDecimal;
import java.text.Normalizer;
import java.util.Locale;
import java.util.regex.Pattern;

/**
 * キー比較用の文字列正規化。
 * Python 版の {@code norm()} と同じく NFKC 正規化 + 空白除去 + 大文字化を行い、
 * 整数値の浮動小数（{@code 72.0}）は {@code "72"} に整数化する。
 */
public final class Norm {

    private static final Pattern WHITESPACE = Pattern.compile("\\s+", Pattern.UNICODE_CHARACTER_CLASS);

    /** 契約NO = ハイフン無し7桁英数字 例: 191352R */
    public static final Pattern KEIYAKU_PATTERN = Pattern.compile("^[0-9A-Z]{7}$");

    private Norm() {
    }

    /** 正規化済みの文字列が契約NO形式（ハイフン無し7桁英数字）か。 */
    public static boolean isKeiyaku(String normalized) {
        return normalized != null && KEIYAKU_PATTERN.matcher(normalized).matches();
    }

    /** NFKC 正規化 + 空白除去 + 大文字化した比較用キーを返す。 */
    public static String norm(Object value) {
        if (value == null) {
            return "";
        }
        String s = Normalizer.normalize(text(value), Normalizer.Form.NFKC);
        return WHITESPACE.matcher(s).replaceAll("").toUpperCase(Locale.ROOT);
    }

    /** 契約NO比較用。正規化した上でハイフンを除去する。 */
    public static String keiyaku(Object value) {
        return norm(value).replace("-", "");
    }

    /** 数値の整数化を含む素の文字列化（正規化はしない）。 */
    public static String text(Object value) {
        if (value == null) {
            return "";
        }
        if (value instanceof Double d) {
            return fromDouble(d);
        }
        if (value instanceof Float f) {
            return fromDouble(f.doubleValue());
        }
        if (value instanceof Number n) {
            return n.toString();
        }
        return String.valueOf(value);
    }

    private static String fromDouble(double d) {
        if (Double.isNaN(d) || Double.isInfinite(d)) {
            return Double.toString(d);
        }
        if (d == Math.rint(d) && Math.abs(d) < 1e15) {
            return Long.toString((long) d);
        }
        return BigDecimal.valueOf(d).stripTrailingZeros().toPlainString();
    }

    /** 数値セルなら {@code double} を、それ以外なら {@code null} を返す。 */
    public static Double number(Object value) {
        return value instanceof Number n ? n.doubleValue() : null;
    }

    /** {@code null}・空文字なら true。 */
    public static boolean isBlank(Object value) {
        return norm(value).isEmpty();
    }
}
