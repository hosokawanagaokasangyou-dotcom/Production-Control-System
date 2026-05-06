package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
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
