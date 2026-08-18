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
 * 加工計画／加工日報の拡張子ポリシー。
 *
 * <p>必須拡張子は環境変数（{@link AppPaths#KEY_PM_AI_PROCESSING_PLAN_REQUIRED_EXT} /
 * {@link AppPaths#KEY_PM_AI_DAILY_REPORT_REQUIRED_EXT}）で上書き可能。フォルダ内の候補のうち更新日時が最新の
 * ファイルが期待拡張子以外ならエラーとし、読込対象は返さない。
 */
public final class SourceFileExtensionPolicy {

    /** @deprecated 既定値。実行時は {@link AppPaths#resolveProcessingPlanRequiredExt(Map)} を使う。 */
    @Deprecated
    public static final String PROCESSING_PLAN_REQUIRED_SUFFIX =
            AppPaths.DEFAULT_PROCESSING_PLAN_REQUIRED_EXT;

    /** @deprecated 既定値。実行時は {@link AppPaths#resolveDailyReportRequiredExt(Map)} を使う。 */
    @Deprecated
    public static final String DAILY_REPORT_REQUIRED_SUFFIX =
            AppPaths.DEFAULT_DAILY_REPORT_REQUIRED_EXT;

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

    /** 既定拡張子（環境変数なし）で加工計画フォルダを検証。 */
    public static Result checkProcessingPlanDirectory(Path dir) {
        return checkProcessingPlanDirectory(dir, Map.of());
    }

    public static Result checkProcessingPlanDirectory(Path dir, Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String required = AppPaths.resolveProcessingPlanRequiredExt(u);
        List<String> candidates = AppPaths.resolveProcessingPlanCandidateExts(u);
        return checkDirectory(
                dir,
                "加工計画",
                required,
                p -> isProcessingPlanCandidate(p, candidates),
                p -> isProcessingPlanLoadable(p, required));
    }

    /** 既定拡張子（環境変数なし）で加工日報フォルダを検証。 */
    public static Result checkDailyReportDirectory(Path dir) {
        return checkDailyReportDirectory(dir, Map.of());
    }

    public static Result checkDailyReportDirectory(Path dir, Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String required = AppPaths.resolveDailyReportRequiredExt(u);
        return checkDirectory(
                dir,
                "加工日報",
                required,
                SourceFileExtensionPolicy::isDailyReportCandidate,
                p -> isDailyReportLoadable(p, required));
    }

    /** 単一ファイル指定時の拡張子検証（加工計画・既定拡張子）。 */
    public static Result checkProcessingPlanFile(Path file) {
        return checkProcessingPlanFile(file, Map.of());
    }

    public static Result checkProcessingPlanFile(Path file, Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        return checkSingleFile(file, "加工計画", AppPaths.resolveProcessingPlanRequiredExt(u));
    }

    /** 単一ファイル指定時の拡張子検証（加工日報・既定拡張子）。 */
    public static Result checkDailyReportFile(Path file) {
        return checkDailyReportFile(file, Map.of());
    }

    public static Result checkDailyReportFile(Path file, Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String required = AppPaths.resolveDailyReportRequiredExt(u);
        if (file == null || !Files.isRegularFile(file)) {
            return Result.emptyDir("加工日報", required);
        }
        Path abs = file.toAbsolutePath().normalize();
        String name = fileName(abs);
        if (!name.startsWith(DAILY_REPORT_FILENAME_PREFIX)) {
            return new Result(
                    false,
                    Optional.empty(),
                    Optional.of(abs),
                    "加工日報のファイル名が不正です（期待接頭辞: "
                            + DAILY_REPORT_FILENAME_PREFIX
                            + "）: "
                            + name);
        }
        if (!endsWithIgnoreCase(name, required)) {
            return Result.mismatch("加工日報", required, abs, extensionOf(name));
        }
        return Result.ok(abs);
    }

    public static boolean isProcessingPlanLoadable(Path p) {
        return isProcessingPlanLoadable(p, AppPaths.DEFAULT_PROCESSING_PLAN_REQUIRED_EXT);
    }

    public static boolean isProcessingPlanLoadable(Path p, String requiredSuffix) {
        return endsWithIgnoreCase(fileName(p), requiredSuffix);
    }

    public static boolean isDailyReportLoadable(Path p) {
        return isDailyReportLoadable(p, AppPaths.DEFAULT_DAILY_REPORT_REQUIRED_EXT);
    }

    public static boolean isDailyReportLoadable(Path p, String requiredSuffix) {
        String name = fileName(p);
        return name.startsWith(DAILY_REPORT_FILENAME_PREFIX)
                && endsWithIgnoreCase(name, requiredSuffix);
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
        for (Result result : blockingMismatchResults(u)) {
            String msg = result.errorMessage();
            if (msg != null && !msg.isBlank()) {
                messages.add(msg);
            }
        }
        return List.copyOf(messages);
    }

    /**
     * ゲート対象の拡張子／ファイル名不一致結果（最新候補パス付き）。空フォルダは含めない。
     */
    public static List<Result> blockingMismatchResults(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        List<Result> out = new ArrayList<>();
        appendMismatchResultIfPresent(out, checkProcessingPlanForUi(u));
        appendMismatchResultIfPresent(out, checkDailyReportForUi(u));
        return List.copyOf(out);
    }

    /**
     * ゲート対象の不正ファイルパス（削除候補）。重複は除く。
     */
    public static List<Path> blockingMismatchPaths(Map<String, String> ui) {
        List<Path> paths = new ArrayList<>();
        for (Result result : blockingMismatchResults(ui)) {
            result.newestCandidatePath()
                    .ifPresent(
                            p -> {
                                Path abs = p.toAbsolutePath().normalize();
                                if (!paths.contains(abs)) {
                                    paths.add(abs);
                                }
                            });
        }
        return List.copyOf(paths);
    }

    private static void appendMismatchResultIfPresent(List<Result> out, Result result) {
        if (result != null && !result.ok() && result.newestCandidatePath().isPresent()) {
            out.add(result);
        }
    }

    /** UI 環境変数から加工計画ソースを検証（明示パス優先）。 */
    public static Result checkProcessingPlanForUi(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String explicit = trim(u.get(AppPaths.KEY_PM_AI_PROCESSING_PLAN_PATH));
        if (!explicit.isEmpty()) {
            Path p = Path.of(explicit);
            if (Files.isRegularFile(p)) {
                return checkProcessingPlanFile(p, u);
            }
        }
        return checkProcessingPlanDirectory(AppPaths.resolveTaskInputSourceDir(u), u);
    }

    /** UI 環境変数から加工日報ソースを検証（明示パス優先）。 */
    public static Result checkDailyReportForUi(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String explicit = trim(u.get("PM_AI_DAILY_REPORT_CSV_PATH"));
        if (!explicit.isEmpty()) {
            Path p = Path.of(explicit);
            if (Files.isRegularFile(p)) {
                return checkDailyReportFile(p, u);
            }
        }
        return checkDailyReportDirectory(AppPaths.resolveDailyReportSourceDir(u), u);
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

    private static boolean isProcessingPlanCandidate(Path p, List<String> candidateSuffixes) {
        String n = fileName(p).toLowerCase(Locale.ROOT);
        for (String s : candidateSuffixes) {
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
