package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.BasicFileAttributes;
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.function.Predicate;
import java.util.stream.Stream;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.NetworkSourceDirResolver;

/**
 * リモートデスクトップ RPA 取得先フォルダから、種別ごとの最新ファイル（フルパス）を解決する。
 */
public final class RemoteDesktopLatestSourceFiles {

    /** 受注明細表 RPA 出力フォルダ（環境マップ上書き可）。 */
    public static final String KEY_ORDER_DETAIL_SOURCE_DIR = AppPaths.KEY_PM_AI_ORDER_DETAIL_SOURCE_DIR;

    private static final String NOT_FOUND = "（該当ファイルなし）";
    private static final String NO_ACQUIRED_AT = "—";

    private static final DateTimeFormatter ACQUIRED_AT_DISPLAY =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss", Locale.JAPAN);

    private static final String ORDER_DETAIL_NAME_MARKER = "受注明細";

    /** 原本転記・計画確認タブ: 加工計画／加工日報の鮮度警告しきい値（1時間）。 */
    public static final long PIPELINE_CHECK_MAX_SOURCE_AGE_MS = 60L * 60L * 1000L;

    public enum Category {
        PROCESSING_PLAN("加工計画"),
        DAILY_REPORT("加工日報"),
        PROCESSING_ACTUALS("加工実績"),
        ORDER_DETAIL("受注明細表");

        private final String label;

        Category(String label) {
            this.label = label;
        }

        public String label() {
            return label;
        }
    }

    /** 共有フォルダ上の最新ファイル1件。 */
    public record ResolvedFile(Path path, long acquiredAtEpochMillis) {
        public String fullPathText() {
            return path.toAbsolutePath().normalize().toString();
        }
    }

    public record Row(Category category, String fullPath, long acquiredAtEpochMillis) {
        public String displayPath() {
            return fullPath != null && !fullPath.isBlank() ? fullPath : NOT_FOUND;
        }

        /** ファイル更新日時（取得日時）の表示。未検出時は {@code —}。 */
        public String displayAcquiredAt() {
            if (fullPath == null || fullPath.isBlank() || acquiredAtEpochMillis <= 0L) {
                return NO_ACQUIRED_AT;
            }
            return ACQUIRED_AT_DISPLAY.format(
                    Instant.ofEpochMilli(acquiredAtEpochMillis).atZone(ZoneId.systemDefault()));
        }
    }

    private RemoteDesktopLatestSourceFiles() {}

    public static List<Row> resolveAll(Map<String, String> ui) {
        Map<String, String> env = ui != null ? ui : Map.of();
        List<Row> rows = new ArrayList<>(Category.values().length);
        for (Category category : Category.values()) {
            Optional<ResolvedFile> resolved = resolveFile(category, env);
            rows.add(
                    new Row(
                            category,
                            resolved.map(ResolvedFile::fullPathText).orElse(""),
                            resolved.map(ResolvedFile::acquiredAtEpochMillis).orElse(0L)));
        }
        return List.copyOf(rows);
    }

    /**
     * 原本転記・計画確認向け: 加工計画・加工日報のうち、{@code maxAgeMillis} より古い取得元のみ返す。
     *
     * <p>ファイル未検出（取得日時なし）は対象外（走査側の警告に委ねる）。
     */
    public static List<Row> pipelineCheckSourcesOlderThan(
            Map<String, String> ui, long maxAgeMillis, long nowEpochMillis) {
        long cutoff = nowEpochMillis - maxAgeMillis;
        return resolveAll(ui).stream()
                .filter(
                        r ->
                                r.category() == Category.PROCESSING_PLAN
                                        || r.category() == Category.DAILY_REPORT)
                .filter(r -> r.acquiredAtEpochMillis() > 0L && r.acquiredAtEpochMillis() < cutoff)
                .toList();
    }

    private static Optional<ResolvedFile> resolveFile(Category category, Map<String, String> ui) {
        return switch (category) {
            case PROCESSING_PLAN ->
                    NetworkSourceDirResolver.newestTaskInputFileInDirectory(
                                    AppPaths.resolveTaskInputSourceDir(ui))
                            .flatMap(RemoteDesktopLatestSourceFiles::toResolvedFile);
            case DAILY_REPORT ->
                    KonanDailyReportLookup.resolveNewestCsvPath(ui)
                            .flatMap(RemoteDesktopLatestSourceFiles::toResolvedFile);
            case PROCESSING_ACTUALS -> resolveNewestActualDetailFile(ui);
            case ORDER_DETAIL -> resolveNewestOrderDetailFile(ui);
        };
    }

    private static Optional<ResolvedFile> resolveNewestActualDetailFile(Map<String, String> ui) {
        String workbook = trim(ui.get(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_WORKBOOK));
        if (!workbook.isEmpty()) {
            Path p = Path.of(workbook).toAbsolutePath().normalize();
            if (Files.isRegularFile(p)) {
                return toResolvedFile(p);
            }
        }
        return NetworkSourceDirResolver.newestExcelFileInDirectory(
                        AppPaths.resolveActualDetailSourceDir(ui))
                .flatMap(RemoteDesktopLatestSourceFiles::toResolvedFile);
    }

    private static Optional<ResolvedFile> resolveNewestOrderDetailFile(Map<String, String> ui) {
        Optional<Path> dirOpt = AppPaths.resolveOrderDetailSourceDir(ui);
        if (dirOpt.isEmpty()) {
            return Optional.empty();
        }
        return pickNewestInDir(
                        dirOpt.get(),
                        p -> {
                            String name =
                                    p.getFileName() != null
                                            ? p.getFileName().toString()
                                            : "";
                            if (name.startsWith("~$")) {
                                return false;
                            }
                            String lower = name.toLowerCase(Locale.ROOT);
                            boolean ext =
                                    lower.endsWith(".xlsx")
                                            || lower.endsWith(".xlsm")
                                            || lower.endsWith(".csv");
                            return ext && name.contains(ORDER_DETAIL_NAME_MARKER);
                        })
                .flatMap(RemoteDesktopLatestSourceFiles::toResolvedFile);
    }

    static Optional<ResolvedFile> toResolvedFile(Path path) {
        if (path == null || !Files.isRegularFile(path)) {
            return Optional.empty();
        }
        long at = acquiredAtEpochMillis(path);
        if (at <= 0L) {
            return Optional.empty();
        }
        return Optional.of(new ResolvedFile(path, at));
    }

    static Optional<Path> pickNewestInDir(Path dir, Predicate<Path> accept) {
        if (!isAccessibleDir(dir)) {
            return Optional.empty();
        }
        try (Stream<Path> stream = Files.list(dir)) {
            return stream.filter(Files::isRegularFile)
                    .filter(accept)
                    .max(Comparator.comparingLong(RemoteDesktopLatestSourceFiles::mtimeScore));
        } catch (IOException e) {
            return Optional.empty();
        }
    }

    private static boolean isAccessibleDir(Path dir) {
        try {
            return Files.isDirectory(dir) && Files.isReadable(dir);
        } catch (Exception e) {
            return false;
        }
    }

    /** ファイル内容の鮮度を表す更新日時（アクセスしただけでは新しく扱わない）。 */
    static long acquiredAtEpochMillis(Path p) {
        return mtimeScore(p);
    }

    private static long mtimeScore(Path p) {
        try {
            BasicFileAttributes a = Files.readAttributes(p, BasicFileAttributes.class);
            return a.lastModifiedTime().toMillis();
        } catch (IOException e) {
            return Long.MIN_VALUE;
        }
    }

    private static String trim(String s) {
        return s != null ? s.strip() : "";
    }
}
