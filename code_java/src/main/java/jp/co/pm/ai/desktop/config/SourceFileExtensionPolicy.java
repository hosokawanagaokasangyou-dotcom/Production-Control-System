package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.function.Predicate;
import java.util.stream.Stream;

/**
 * 加工計画（xlsx）／加工日報（csv）の拡張子ポリシー。
 *
 * <p>フォルダ内の候補のうち更新日時が最新のファイルが期待拡張子以外ならエラーとし、読込対象は返さない。
 */
public final class SourceFileExtensionPolicy {

    /** 加工計画フォルダで「最新判定」に含める拡張子（読込は {@link #PROCESSING_PLAN_REQUIRED_SUFFIX} のみ）。 */
    private static final List<String> PROCESSING_PLAN_CANDIDATE_SUFFIXES =
            List.of(".csv", ".parquet", ".pq", ".xlsx", ".xlsm", ".xltx", ".xltm", ".xls");

    public static final String PROCESSING_PLAN_REQUIRED_SUFFIX = ".xlsx";

    public static final String DAILY_REPORT_REQUIRED_SUFFIX = ".csv";

    public static final String DAILY_REPORT_FILENAME_PREFIX = "加工日報発行問合せ_";

    public record Result(
            boolean ok, Optional<Path> loadablePath, Optional<Path> newestCandidatePath, String errorMessage) {

        public static Result ok(Path loadable) {
            Path abs = loadable.toAbsolutePath().normalize();
            return new Result(true, Optional.of(abs), Optional.of(abs), "");
        }

        public static Result emptyDir(String label, String requiredSuffix) {
            return new Result(
                    false,
                    Optional.empty(),
                    Optional.empty(),
                    label + "の読込対象（*" + requiredSuffix + "）が見つかりません。");
        }

        public static Result mismatch(
                String label, String requiredSuffix, Path newest, String actualSuffix) {
            String name =
                    newest.getFileName() != null ? newest.getFileName().toString() : newest.toString();
            return new Result(
                    false,
                    Optional.empty(),
                    Optional.of(newest.toAbsolutePath().normalize()),
                    label
                            + "の最新ファイル拡張子が不正です（期待: "
                            + requiredSuffix
                            + " / 実際: "
                            + (actualSuffix.isEmpty() ? "（拡張子なし）" : actualSuffix)
                            + "）: "
                            + name);
        }
    }

    private SourceFileExtensionPolicy() {}

    public static Result checkProcessingPlanDirectory(Path dir) {
        return checkDirectory(
                dir,
                "加工計画",
                PROCESSING_PLAN_REQUIRED_SUFFIX,
                SourceFileExtensionPolicy::isProcessingPlanCandidate,
                SourceFileExtensionPolicy::isProcessingPlanLoadable);
    }

    public static Result checkDailyReportDirectory(Path dir) {
        return checkDirectory(
                dir,
                "加工日報",
                DAILY_REPORT_REQUIRED_SUFFIX,
                SourceFileExtensionPolicy::isDailyReportCandidate,
                SourceFileExtensionPolicy::isDailyReportLoadable);
    }

    /** 単一ファイル指定時の拡張子検証（加工計画）。 */
    public static Result checkProcessingPlanFile(Path file) {
        return checkSingleFile(file, "加工計画", PROCESSING_PLAN_REQUIRED_SUFFIX);
    }

    /** 単一ファイル指定時の拡張子検証（加工日報）。 */
    public static Result checkDailyReportFile(Path file) {
        return checkSingleFile(file, "加工日報", DAILY_REPORT_REQUIRED_SUFFIX);
    }

    public static boolean isProcessingPlanLoadable(Path p) {
        return endsWithIgnoreCase(fileName(p), PROCESSING_PLAN_REQUIRED_SUFFIX);
    }

    public static boolean isDailyReportLoadable(Path p) {
        String name = fileName(p);
        return name.startsWith(DAILY_REPORT_FILENAME_PREFIX)
                && endsWithIgnoreCase(name, DAILY_REPORT_REQUIRED_SUFFIX);
    }

    /**
     * 段階1実行前ゲート用: 最新ファイルの拡張子不一致メッセージのみ返す（空フォルダは含めない）。
     *
     * <p>加工計画は {@link AppPaths#KEY_PM_AI_PROCESSING_PLAN_PATH} があれば単一ファイル、なければ
     * {@link AppPaths#resolveTaskInputSourceDir}。加工日報は {@code PM_AI_DAILY_REPORT_CSV_PATH} があれば単一、
     * なければ {@link AppPaths#resolveDailyReportSourceDir}。
     */
    public static List<String> blockingMismatchMessages(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        List<String> messages = new ArrayList<>();
        appendMismatchIfPresent(messages, checkProcessingPlanForUi(u));
        appendMismatchIfPresent(messages, checkDailyReportForUi(u));
        return List.copyOf(messages);
    }

    private static void appendMismatchIfPresent(List<String> messages, Result result) {
        if (result != null && !result.ok() && result.newestCandidatePath().isPresent()) {
            String msg = result.errorMessage();
            if (msg != null && !msg.isBlank()) {
                messages.add(msg);
            }
        }
    }

    private static Result checkProcessingPlanForUi(Map<String, String> ui) {
        String explicit = trim(ui.get(AppPaths.KEY_PM_AI_PROCESSING_PLAN_PATH));
        if (!explicit.isEmpty()) {
            Path p = Path.of(explicit);
            if (Files.isRegularFile(p)) {
                return checkProcessingPlanFile(p);
            }
        }
        return checkProcessingPlanDirectory(AppPaths.resolveTaskInputSourceDir(ui));
    }

    private static Result checkDailyReportForUi(Map<String, String> ui) {
        String explicit = trim(ui.get("PM_AI_DAILY_REPORT_CSV_PATH"));
        if (!explicit.isEmpty()) {
            Path p = Path.of(explicit);
            if (Files.isRegularFile(p)) {
                return checkDailyReportFile(p);
            }
        }
        return checkDailyReportDirectory(AppPaths.resolveDailyReportSourceDir(ui));
    }

    private static String trim(String s) {
        return s != null ? s.strip() : "";
    }

    private static Result checkSingleFile(Path file, String label, String requiredSuffix) {
        if (file == null || !Files.isRegularFile(file)) {
            return Result.emptyDir(label, requiredSuffix);
        }
        Path abs = file.toAbsolutePath().normalize();
        String suffix = extensionOf(fileName(abs));
        if (!endsWithIgnoreCase(fileName(abs), requiredSuffix)) {
            return Result.mismatch(label, requiredSuffix, abs, suffix);
        }
        return Result.ok(abs);
    }

    private static Result checkDirectory(
            Path dir,
            String label,
            String requiredSuffix,
            Predicate<Path> candidate,
            Predicate<Path> loadable) {
        if (dir == null || !isAccessibleDir(dir)) {
            return Result.emptyDir(label, requiredSuffix);
        }
        Optional<Path> newest = pickNewest(dir, candidate);
        if (newest.isEmpty()) {
            return Result.emptyDir(label, requiredSuffix);
        }
        Path newestPath = newest.get();
        if (!loadable.test(newestPath)) {
            return Result.mismatch(label, requiredSuffix, newestPath, extensionOf(fileName(newestPath)));
        }
        return Result.ok(newestPath);
    }

    private static Optional<Path> pickNewest(Path dir, Predicate<Path> accept) {
        try (Stream<Path> stream = Files.list(dir)) {
            return stream.filter(Files::isRegularFile)
                    .filter(p -> !isLockFile(p))
                    .filter(accept)
                    .max(Comparator.comparingLong(SourceFileExtensionPolicy::mtimeScore));
        } catch (IOException e) {
            return Optional.empty();
        }
    }

    private static boolean isProcessingPlanCandidate(Path p) {
        String n = fileName(p).toLowerCase(Locale.ROOT);
        for (String s : PROCESSING_PLAN_CANDIDATE_SUFFIXES) {
            if (n.endsWith(s)) {
                return true;
            }
        }
        return false;
    }

    private static boolean isDailyReportCandidate(Path p) {
        return fileName(p).startsWith(DAILY_REPORT_FILENAME_PREFIX);
    }

    private static boolean isAccessibleDir(Path dir) {
        try {
            return Files.isDirectory(dir) && Files.isReadable(dir);
        } catch (Exception e) {
            return false;
        }
    }

    private static boolean isLockFile(Path p) {
        return fileName(p).startsWith("~$");
    }

    private static String fileName(Path p) {
        return p != null && p.getFileName() != null ? p.getFileName().toString() : "";
    }

    private static boolean endsWithIgnoreCase(String name, String suffix) {
        return name.toLowerCase(Locale.ROOT).endsWith(suffix.toLowerCase(Locale.ROOT));
    }

    private static String extensionOf(String fileName) {
        int dot = fileName.lastIndexOf('.');
        return dot >= 0 ? fileName.substring(dot).toLowerCase(Locale.ROOT) : "";
    }

    private static long mtimeScore(Path p) {
        try {
            BasicFileAttributes a = Files.readAttributes(p, BasicFileAttributes.class);
            return a.lastModifiedTime().toMillis();
        } catch (IOException e) {
            return Long.MIN_VALUE;
        }
    }
}
