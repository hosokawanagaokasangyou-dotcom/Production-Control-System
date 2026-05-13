package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Locale;
import java.util.Objects;
import java.util.function.Predicate;
import java.util.stream.Stream;

/**
 * 段階2の output 成果物ファイル名（Python {@code stage2_output_naming} と整合）。
 * 本体ブックは「計画」「人員」の日本語接頭辞＋16桁の数字のみ（サイドカーの表・論・一覧などは除外）。
 */
public final class Stage2OutputNaming {

    public static final String PLAN_PREFIX = "計画";
    public static final String MEMBER_PREFIX = "人員";
    /** yyMMddHHmmss（12）＋下4桁 */
    public static final int STAMP_DIGITS = 16;

    private Stage2OutputNaming() {}

    /**
     * Python {@code format_stage2_stamp} と同形の 16 桁（yyMMddHHmmss ＋実行時刻の下位 4 桁。Java は {@link
     * LocalDateTime#getNano} をマイクロ秒相当に縮約）。
     */
    public static String formatStage2Stamp(LocalDateTime baseWallClock, LocalDateTime runWallClock) {
        DateTimeFormatter coreFmt = DateTimeFormatter.ofPattern("yyMMddHHmmss", Locale.ROOT);
        String core = baseWallClock.format(coreFmt);
        int frac = (runWallClock.getNano() / 1000) % 10000;
        return core + String.format(Locale.ROOT, "%04d", frac);
    }

    static boolean isDigitStem(String s) {
        if (s == null || s.length() != STAMP_DIGITS) {
            return false;
        }
        return s.chars().allMatch(Character::isDigit);
    }

    private static boolean matchesPrimary(Path path, String prefix, String ext) {
        Path fn = path.getFileName();
        if (fn == null) {
            return false;
        }
        String n = fn.toString();
        if (!n.endsWith(ext)) {
            return false;
        }
        String stem = n.substring(0, n.length() - ext.length());
        if (!stem.startsWith(prefix)) {
            return false;
        }
        if (stem.length() != prefix.length() + STAMP_DIGITS) {
            return false;
        }
        return isDigitStem(stem.substring(prefix.length()));
    }

    public static boolean isLegacyPlanXlsx(Path path) {
        Path fn = path.getFileName();
        if (fn == null) {
            return false;
        }
        String n = fn.toString();
        return n.startsWith("production_plan_multi_day_") && n.endsWith(".xlsx");
    }

    public static boolean isLegacyPlanJson(Path path) {
        Path fn = path.getFileName();
        if (fn == null) {
            return false;
        }
        String n = fn.toString();
        if (!n.startsWith("production_plan_multi_day_") || !n.endsWith(".json")) {
            return false;
        }
        String stem = n.substring(0, n.length() - 5);
        return !stem.endsWith("_tabular_source")
                && !stem.endsWith("_logical_view")
                && !stem.endsWith("_equipment_gantt_contract")
                && !stem.endsWith("_actual_detail_gantt_contract")
                && !stem.endsWith("_結果_タスク一覧");
    }

    public static boolean isLegacyMemberXlsx(Path path) {
        Path fn = path.getFileName();
        if (fn == null) {
            return false;
        }
        String n = fn.toString();
        return n.startsWith("member_schedule_") && n.endsWith(".xlsx");
    }

    public static boolean isLegacyMemberJson(Path path) {
        Path fn = path.getFileName();
        if (fn == null) {
            return false;
        }
        String n = fn.toString();
        return n.startsWith("member_schedule_") && n.endsWith(".json");
    }

    /** ディレクトリ内で条件に合うもののうち最新（更新時刻最大）。 */
    public static Path newestMatching(Path dir, Predicate<Path> accept) throws IOException {
        Objects.requireNonNull(accept);
        if (!Files.isDirectory(dir)) {
            return null;
        }
        Path best = null;
        long bestTime = Long.MIN_VALUE;
        try (Stream<Path> stream = Files.list(dir)) {
            java.util.Iterator<Path> it = stream.iterator();
            while (it.hasNext()) {
                Path p = it.next();
                if (!Files.isRegularFile(p) || !accept.test(p)) {
                    continue;
                }
                long t = Files.getLastModifiedTime(p).toMillis();
                if (t >= bestTime) {
                    bestTime = t;
                    best = p;
                }
            }
        }
        return best;
    }

    /** 新命名または旧命名の計画ブック（xlsx）のうち最新。 */
    public static Path newestPrimaryPlanXlsx(Path dir) throws IOException {
        Path a = newestMatching(dir, p -> matchesPrimary(p, PLAN_PREFIX, ".xlsx"));
        Path b = newestMatching(dir, Stage2OutputNaming::isLegacyPlanXlsx);
        return newerOf(a, b);
    }

    /** 新命名または旧命名の計画ブック JSON ミラーのうち最新。 */
    public static Path newestPrimaryPlanJson(Path dir) throws IOException {
        Path a = newestMatching(dir, p -> matchesPrimary(p, PLAN_PREFIX, ".json"));
        Path b = newestMatching(dir, Stage2OutputNaming::isLegacyPlanJson);
        return newerOf(a, b);
    }

    /** 計画系 primary JSON（新命名・旧命名）か。 */
    public static boolean acceptsPrimaryPlanJson(Path path) {
        return matchesPrimary(path, PLAN_PREFIX, ".json") || isLegacyPlanJson(path);
    }

    /**
     * ディレクトリ内の計画 primary JSON の {@link Files#getLastModifiedTime} の最大値。該当なし・ディレクトリ不正時は
     * {@code 0}。
     */
    public static long maxPrimaryPlanJsonLastModifiedMillis(Path dir) throws IOException {
        if (dir == null || !Files.isDirectory(dir)) {
            return 0L;
        }
        long max = 0L;
        try (Stream<Path> stream = Files.list(dir)) {
            java.util.Iterator<Path> it = stream.iterator();
            while (it.hasNext()) {
                Path p = it.next();
                if (!Files.isRegularFile(p) || !acceptsPrimaryPlanJson(p)) {
                    continue;
                }
                max = Math.max(max, Files.getLastModifiedTime(p).toMillis());
            }
        }
        return max;
    }

    /**
     * {@code lastModified} が {@code strictlyAfterMillis} より大きい計画 primary JSON のうち最新（同時刻はパス文字列で
     * 後勝ち）。無ければ {@code null}。
     */
    public static Path newestPrimaryPlanJsonAfter(Path dir, long strictlyAfterMillis) throws IOException {
        if (dir == null || !Files.isDirectory(dir)) {
            return null;
        }
        Path best = null;
        long bestTime = Long.MIN_VALUE;
        try (Stream<Path> stream = Files.list(dir)) {
            java.util.Iterator<Path> it = stream.iterator();
            while (it.hasNext()) {
                Path p = it.next();
                if (!Files.isRegularFile(p) || !acceptsPrimaryPlanJson(p)) {
                    continue;
                }
                long t = Files.getLastModifiedTime(p).toMillis();
                if (t <= strictlyAfterMillis) {
                    continue;
                }
                if (t > bestTime || (t == bestTime && best != null && p.toString().compareTo(best.toString()) > 0)) {
                    bestTime = t;
                    best = p;
                }
            }
        }
        return best;
    }

    /** 新命名または旧命名の個人別ブック（xlsx）のうち最新。 */
    public static Path newestPrimaryMemberXlsx(Path dir) throws IOException {
        Path a = newestMatching(dir, p -> matchesPrimary(p, MEMBER_PREFIX, ".xlsx"));
        Path b = newestMatching(dir, Stage2OutputNaming::isLegacyMemberXlsx);
        return newerOf(a, b);
    }

    /** 新命名または旧命名の個人別 JSON のうち最新。 */
    public static Path newestPrimaryMemberJson(Path dir) throws IOException {
        Path a = newestMatching(dir, p -> matchesPrimary(p, MEMBER_PREFIX, ".json"));
        Path b = newestMatching(dir, Stage2OutputNaming::isLegacyMemberJson);
        return newerOf(a, b);
    }

    private static Path newerOf(Path a, Path b) throws IOException {
        if (a == null) {
            return b;
        }
        if (b == null) {
            return a;
        }
        long ta = Files.getLastModifiedTime(a).toMillis();
        long tb = Files.getLastModifiedTime(b).toMillis();
        return ta >= tb ? a : b;
    }
}
