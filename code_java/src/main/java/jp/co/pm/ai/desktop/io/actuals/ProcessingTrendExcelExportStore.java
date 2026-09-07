package jp.co.pm.ai.desktop.io.actuals;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Stream;

import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * 加工トレンド Excel の固定出力先（{@code .pm-ai-cache/exports/processing-trend/}）。
 * 最新1件のみ残す運用向けのパス解決・掃除・最新検出。
 */
public final class ProcessingTrendExcelExportStore {

    private ProcessingTrendExcelExportStore() {}

    public static Path resolveDirectory(Map<String, String> ui) {
        return AppPaths.resolveRepoRoot(ui != null ? ui : Map.of())
                .resolve(".pm-ai-cache")
                .resolve("exports")
                .resolve("processing-trend")
                .toAbsolutePath()
                .normalize();
    }

    /** ディレクトリを用意し、既存 {@code *.xlsx} を消してから出力パスを返す（ファイルはまだ作らない）。 */
    public static Path prepareTarget(Path dir, String fileName) throws IOException {
        if (dir == null) {
            throw new IllegalArgumentException("dir");
        }
        String name = fileName != null ? fileName.strip() : "";
        if (name.isEmpty()) {
            throw new IllegalArgumentException("fileName");
        }
        Files.createDirectories(dir);
        deleteAllXlsx(dir);
        return dir.resolve(name).toAbsolutePath().normalize();
    }

    public static void deleteAllXlsx(Path dir) throws IOException {
        if (dir == null || !Files.isDirectory(dir)) {
            return;
        }
        try (Stream<Path> stream = Files.list(dir)) {
            for (Path path : stream.toList()) {
                if (!Files.isRegularFile(path)) {
                    continue;
                }
                String n = path.getFileName() != null ? path.getFileName().toString() : "";
                if (n.toLowerCase(Locale.ROOT).endsWith(".xlsx")) {
                    Files.deleteIfExists(path);
                }
            }
        }
    }

    public static Optional<Path> findNewestXlsx(Path dir) {
        if (dir == null || !Files.isDirectory(dir)) {
            return Optional.empty();
        }
        try (Stream<Path> stream = Files.list(dir)) {
            return stream
                    .filter(Files::isRegularFile)
                    .filter(
                            p -> {
                                String n = p.getFileName() != null ? p.getFileName().toString() : "";
                                return n.toLowerCase(Locale.ROOT).endsWith(".xlsx");
                            })
                    .max(
                            Comparator.comparingLong(ProcessingTrendExcelExportStore::lastModifiedMillis)
                                    .thenComparing(p -> p.getFileName().toString()));
        } catch (IOException ex) {
            return Optional.empty();
        }
    }

    private static long lastModifiedMillis(Path p) {
        try {
            return Files.getLastModifiedTime(p).toMillis();
        } catch (IOException e) {
            return Long.MIN_VALUE;
        }
    }
}
