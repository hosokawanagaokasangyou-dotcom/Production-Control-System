package jp.co.pm.ai.kouchin.verify;

import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 対象年月（年・月）。
 * 「2026年7月度」「RVSHEET202607」「対象年月 : 2026年07月」の各表記から生成する。
 */
public record YearMonthKey(int year, int month) implements Comparable<YearMonthKey> {

    /** 「yyyy年m月度」 */
    public static final Pattern GATSUDO_PATTERN = Pattern.compile("(\\d{4})\\s*年\\s*(\\d{1,2})\\s*月度");
    /** 「yyyy年m月」 */
    public static final Pattern YEAR_MONTH_PATTERN = Pattern.compile("(\\d{4})\\s*年\\s*(\\d{1,2})\\s*月");
    /** 「RVSHEETyyyymm」 */
    public static final Pattern RVSHEET_PATTERN = Pattern.compile("RVSHEET(\\d{4})(\\d{2})", Pattern.CASE_INSENSITIVE);
    /** 「2026-08」「2026/8」「2026.08」 */
    public static final Pattern DASHED_PATTERN = Pattern.compile("^(\\d{4})[-/.](\\d{1,2})$");
    /** 「202608」 */
    public static final Pattern YYYYMM_PATTERN = Pattern.compile("^(\\d{4})(\\d{2})$");

    public YearMonthKey {
        if (month < 1 || month > 12) {
            throw new IllegalArgumentException("月は1〜12で指定してください: " + month);
        }
    }

    /** 「yyyy年m月度」を含む文字列から生成する。 */
    public static Optional<YearMonthKey> parseGatsudo(String source) {
        return find(GATSUDO_PATTERN, source);
    }

    /** 「yyyy年m月」を含む文字列から生成する（「対象年月 : 2026年07月」など）。 */
    public static Optional<YearMonthKey> parseYearMonth(String source) {
        return find(YEAR_MONTH_PATTERN, source);
    }

    /** 「RVSHEETyyyymm」を含む文字列（ファイル名）から生成する。 */
    public static Optional<YearMonthKey> parseRvsheet(String source) {
        return find(RVSHEET_PATTERN, source);
    }

    /**
     * 手動判定.csv / 前月過不足.csv の対象月。
     * 「2026年8月度」「2026-08」「2026/8」「202608」を受け付ける。
     */
    public static Optional<YearMonthKey> parseFlexible(String source) {
        Optional<YearMonthKey> g = parseGatsudo(source);
        if (g.isPresent()) {
            return g;
        }
        Optional<YearMonthKey> ym = parseYearMonth(source);
        if (ym.isPresent()) {
            return ym;
        }
        if (source == null || source.isBlank()) {
            return Optional.empty();
        }
        String n = Norm.norm(source);
        return findFull(DASHED_PATTERN, n).or(() -> findFull(YYYYMM_PATTERN, n));
    }

    private static Optional<YearMonthKey> findFull(Pattern pattern, String source) {
        Matcher m = pattern.matcher(source);
        if (!m.matches()) {
            return Optional.empty();
        }
        int y = Integer.parseInt(m.group(1));
        int mo = Integer.parseInt(m.group(2));
        if (mo < 1 || mo > 12) {
            return Optional.empty();
        }
        return Optional.of(new YearMonthKey(y, mo));
    }

    private static Optional<YearMonthKey> find(Pattern pattern, String source) {
        if (source == null || source.isEmpty()) {
            return Optional.empty();
        }
        Matcher m = pattern.matcher(Norm.norm(source));
        if (!m.find()) {
            return Optional.empty();
        }
        int y = Integer.parseInt(m.group(1));
        int mo = Integer.parseInt(m.group(2));
        if (mo < 1 || mo > 12) {
            return Optional.empty();
        }
        return Optional.of(new YearMonthKey(y, mo));
    }

    /** 年月を通し番号にした値（月差の計算用）。 */
    public int monthIndex() {
        return year * 12 + month;
    }

    /** n か月後（負値で n か月前）。 */
    public YearMonthKey plusMonths(int months) {
        int total = year * 12 + (month - 1) + months;
        return new YearMonthKey(Math.floorDiv(total, 12), Math.floorMod(total, 12) + 1);
    }

    /** 「yyyy年m月度」 */
    public String gatsudoLabel() {
        return year + "年" + month + "月度";
    }

    /** 「yyyy年m月」 */
    public String ymLabel() {
        return year + "年" + month + "月";
    }

    /** 「m月」 */
    public String monthLabel() {
        return month + "月";
    }

    @Override
    public int compareTo(YearMonthKey other) {
        return Integer.compare(monthIndex(), other.monthIndex());
    }

    @Override
    public String toString() {
        return gatsudoLabel();
    }
}
