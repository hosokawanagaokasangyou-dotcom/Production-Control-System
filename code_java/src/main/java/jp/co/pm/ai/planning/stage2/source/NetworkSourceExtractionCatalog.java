package jp.co.pm.ai.planning.stage2.source;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Stream;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.NetworkSourceDirResolver;
import jp.co.pm.ai.desktop.reconciliation.KonanDailyReportLookup;
import jp.co.pm.ai.planning.stage2.Stage2PlanRunDateResolver;

/** 加工計画・加工日報フォルダを走査し、取得時刻付きエントリ一覧を構築する。 */
public final class NetworkSourceExtractionCatalog {

    public record SourceEntry(
            Path path,
            LocalDateTime extractionTime,
            NetworkSourceExtractionTimeSupport.SourceKind timeSourceKind,
            String fileName) {}

    private NetworkSourceExtractionCatalog() {}

    public static List<SourceEntry> listProcessingPlans(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path dir = AppPaths.resolveTaskInputSourceDir(u);
        if (!isListableDir(dir)) {
            return List.of();
        }
        List<SourceEntry> out = new ArrayList<>();
        try (Stream<Path> stream = Files.list(dir)) {
            for (Path p : stream.toList()) {
                if (!isPlanCandidate(p)) {
                    continue;
                }
                resolvePlanExtractionTime(p).ifPresent(dt -> out.add(toEntry(p, dt)));
            }
        } catch (IOException ex) {
            return List.of();
        }
        out.sort(Comparator.comparing(SourceEntry::extractionTime).reversed());
        return List.copyOf(out);
    }

    public static List<SourceEntry> listDailyReports(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String dirText = u.getOrDefault(KonanDailyReportLookup.KEY_DAILY_REPORT_SOURCE_DIR, "").strip();
        if (dirText.isEmpty()) {
            dirText = "\\\\192.168.0.101\\共有フォルダ\\湖南工場\\湖南共有\\生産管理システム\\管理システム\\●DATA\\加工日報";
        }
        Path dir = Path.of(dirText);
        if (!isListableDir(dir)) {
            return List.of();
        }
        List<SourceEntry> out = new ArrayList<>();
        try (Stream<Path> stream = Files.list(dir)) {
            for (Path p : stream.toList()) {
                if (!isDailyReportCandidate(p)) {
                    continue;
                }
                resolveDailyReportExtractionTime(p).ifPresent(dt -> out.add(toEntry(p, dt)));
            }
        } catch (IOException ex) {
            return List.of();
        }
        out.sort(Comparator.comparing(SourceEntry::extractionTime).reversed());
        return List.copyOf(out);
    }

    public static Optional<SourceEntry> resolvePlanEntry(Path path) {
        return resolvePlanExtractionTime(path).map(dt -> toEntry(path, dt));
    }

    public static Optional<SourceEntry> resolveDailyReportEntry(Path path) {
        return resolveDailyReportExtractionTime(path).map(dt -> toEntry(path, dt));
    }

    private static SourceEntry toEntry(
            Path path, ResolvedPlanTime resolved) {
        String name = path.getFileName() != null ? path.getFileName().toString() : path.toString();
        return new SourceEntry(
                path.toAbsolutePath().normalize(),
                resolved.dateTime(),
                resolved.kind(),
                name);
    }

    private record ResolvedPlanTime(
            LocalDateTime dateTime, NetworkSourceExtractionTimeSupport.SourceKind kind) {}

    private static Optional<ResolvedPlanTime> resolvePlanExtractionTime(Path path) {
        Optional<LocalDateTime> sheet =
                Stage2PlanRunDateResolver.readExtractionDateTimeFromWorkbook(path);
        if (sheet.isPresent()) {
            return Optional.of(
                    new ResolvedPlanTime(
                            sheet.get(),
                            NetworkSourceExtractionTimeSupport.SourceKind.SHEET_COLUMN));
        }
        String name = path.getFileName() != null ? path.getFileName().toString() : "";
        Optional<LocalDateTime> fromName = NetworkSourceExtractionTimeSupport.parseFilenameTimestamp(name);
        if (fromName.isPresent()) {
            return Optional.of(
                    new ResolvedPlanTime(
                            fromName.get(),
                            NetworkSourceExtractionTimeSupport.SourceKind.FILENAME));
        }
        return NetworkSourceExtractionTimeSupport.fromFileMtime(path)
                .map(
                        dt ->
                                new ResolvedPlanTime(
                                        dt, NetworkSourceExtractionTimeSupport.SourceKind.FILE_MTIME));
    }

    private static Optional<ResolvedPlanTime> resolveDailyReportExtractionTime(Path path) {
        String name = path.getFileName() != null ? path.getFileName().toString() : "";
        Optional<LocalDateTime> fromName =
                NetworkSourceExtractionTimeSupport.parseFilenameTimestamp(name);
        if (fromName.isPresent()) {
            return Optional.of(
                    new ResolvedPlanTime(
                            fromName.get(),
                            NetworkSourceExtractionTimeSupport.SourceKind.FILENAME));
        }
        Optional<LocalDateTime> meta = NetworkSourceExtractionTimeSupport.parseDailyReportCsvMeta(path);
        if (meta.isPresent()) {
            return Optional.of(
                    new ResolvedPlanTime(meta.get(), NetworkSourceExtractionTimeSupport.SourceKind.CSV_META));
        }
        return NetworkSourceExtractionTimeSupport.fromFileMtime(path)
                .map(
                        dt ->
                                new ResolvedPlanTime(
                                        dt, NetworkSourceExtractionTimeSupport.SourceKind.FILE_MTIME));
    }

    private static boolean isPlanCandidate(Path p) {
        if (p == null || !Files.isRegularFile(p)) {
            return false;
        }
        String name = p.getFileName() != null ? p.getFileName().toString() : "";
        if (name.startsWith("~$")) {
            return false;
        }
        String lower = name.toLowerCase(Locale.ROOT);
        return lower.endsWith(".xlsx")
                || lower.endsWith(".xlsm")
                || lower.endsWith(".csv")
                || lower.endsWith(".parquet")
                || lower.endsWith(".pq");
    }

    private static boolean isDailyReportCandidate(Path p) {
        if (p == null || !Files.isRegularFile(p)) {
            return false;
        }
        String name = p.getFileName() != null ? p.getFileName().toString() : "";
        return name.startsWith("加工日報発行問合せ_")
                && name.toLowerCase(Locale.ROOT).endsWith(".csv");
    }

    private static boolean isListableDir(Path dir) {
        try {
            return dir != null && Files.isDirectory(dir) && Files.isReadable(dir);
        } catch (Exception ex) {
            return false;
        }
    }
}
