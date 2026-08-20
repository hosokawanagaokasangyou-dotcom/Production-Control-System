package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.Stream;

/**
 * 配台マスタ JSON / master ブックの保存直前世代バックアップ。
 * 正本と同階層の {@code {stem}-{ext}-history} にコピーし、古い順に上限まで残す。
 */
public final class MasterDispatchSheetsGenerationBackup {

    public static final int MAX_GENERATIONS = 20;

    private static final ZoneId TOKYO = ZoneId.of("Asia/Tokyo");

    private static final DateTimeFormatter TS =
            DateTimeFormatter.ofPattern("yyyyMMdd-HHmmssSSS").withZone(TOKYO);

    private MasterDispatchSheetsGenerationBackup() {}

    public static Path historyDirFor(Path sourceFile) {
        Path src = Objects.requireNonNull(sourceFile, "sourceFile").toAbsolutePath().normalize();
        Path parent = src.getParent();
        if (parent == null) {
            parent = src.getFileSystem().getPath(".");
        }
        String name = src.getFileName() != null ? src.getFileName().toString() : "file";
        int dot = name.lastIndexOf('.');
        String stem = dot > 0 ? name.substring(0, dot) : name;
        String ext = dot > 0 ? name.substring(dot + 1) : "dat";
        return parent.resolve(stem + "-" + ext.toLowerCase(Locale.ROOT) + "-history")
                .toAbsolutePath()
                .normalize();
    }

    /**
     * 正本が存在するとき履歴へコピーする。無いときは empty（履歴フォルダも作らない）。
     */
    public static Optional<Path> backupExisting(Path sourceFile) throws IOException {
        if (sourceFile == null || !Files.isRegularFile(sourceFile)) {
            return Optional.empty();
        }
        Path src = sourceFile.toAbsolutePath().normalize();
        Path history = historyDirFor(src);
        Files.createDirectories(history);
        String originalName = src.getFileName().toString();
        Path dest = uniqueBackupPath(history, originalName);
        Files.copy(src, dest, StandardCopyOption.REPLACE_EXISTING);
        prune(history);
        return Optional.of(dest);
    }

    static void prune(Path historyDir) throws IOException {
        if (historyDir == null || !Files.isDirectory(historyDir)) {
            return;
        }
        List<Path> files;
        try (Stream<Path> stream = Files.list(historyDir)) {
            files =
                    stream.filter(Files::isRegularFile)
                            .sorted(
                                    Comparator.comparingLong(
                                                    MasterDispatchSheetsGenerationBackup::lastModifiedMillis)
                                            .thenComparing(p -> p.getFileName().toString()))
                            .collect(java.util.stream.Collectors.toCollection(ArrayList::new));
        }
        while (files.size() > MAX_GENERATIONS) {
            Files.deleteIfExists(files.remove(0));
        }
    }

    private static Path uniqueBackupPath(Path history, String originalName) throws IOException {
        String stamp = TS.format(ZonedDateTime.now(TOKYO));
        Path dest = history.resolve(stamp + "_" + originalName);
        int n = 2;
        while (Files.exists(dest)) {
            dest = history.resolve(stamp + "-" + n + "_" + originalName);
            n++;
        }
        return dest;
    }

    private static long lastModifiedMillis(Path p) {
        try {
            return Files.getLastModifiedTime(p).toMillis();
        } catch (IOException e) {
            return Long.MAX_VALUE;
        }
    }
}
