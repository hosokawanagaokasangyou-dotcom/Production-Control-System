package jp.co.pm.ai.kouchin.trend;

import java.text.Normalizer;
import java.util.Set;
import java.util.regex.Pattern;

import jp.co.pm.ai.kouchin.verify.Norm;

/** 工程名・見出し用の NFKC 正規化（大文字化しない）。 */
final class TrendText {

    static final Set<String> TOTAL_LABELS =
            Set.of("加工量合計", "加工賃合計", "営業入庫量", "東レ合計", "総合計");

    private static final Pattern WS = Pattern.compile("\\s+");

    private TrendText() {}

    static String nfkc(Object value) {
        if (value == null) {
            return "";
        }
        return Normalizer.normalize(Norm.text(value), Normalizer.Form.NFKC).strip();
    }

    static String disp(Object value) {
        return nfkc(value).replace('\n', ' ').replace('\r', ' ').strip();
    }

    static String processName(Object value) {
        String s = disp(value);
        if (s.isEmpty()) {
            return "";
        }
        return WS.matcher(s).replaceAll(" ");
    }

    static double num(Object value) {
        Double n = Norm.number(value);
        return n == null ? 0.0 : n;
    }

    static long roundDisp(double value) {
        if (value >= 0) {
            return (long) (value + 0.5);
        }
        return -((long) (-value + 0.5));
    }
}
