package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.DirectoryStream;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.StandardCopyOption;
import java.nio.file.attribute.BasicFileAttributes;
import java.time.OffsetDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import com.fasterxml.jackson.databind.ObjectMapper;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.OperatorUserPaths;

/**
 * 同一化チェック結果（配台計画 Excel ＋ 加工計画 JSON）の操作者別履歴。
 * 共有 DATA 配下にセット保存し、操作者あたり最新 {@link #MAX_SNAPSHOTS_PER_USER} 件を保持する。
 */
public final class IdentityCheckHistoryStore {

    public static final int MAX_SNAPSHOTS_PER_USER = 20;

    public static final String EXCEL_FILE = "配台計画.xlsx";

    public static final String PLAN_JSON_FILE = "加工計画.json";

    public static final String META_FILE = "meta.json";

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final DateTimeFormatter FOLDER_TS =
            DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss");

    private IdentityCheckHistoryStore() {}

    public record Meta(
            String savedAt,
            String operator,
            String result,
            String badgeText,
            int diffCount,
            String excelSourcePath,
            String planSourcePath,
            String excelFileName,
            String planJsonFileName) {}

    public record SnapshotRef(Path dir, Meta meta) {}

    public static Path resolveRoot(Map<String, String> ui) {
        return AppPaths.resolveIdentityCheckHistoryRoot(ui);
    }

    public static Path resolveOperatorDir(Map<String, String> ui, String operator) {
        Path root = resolveRoot(ui).toAbsolutePath().normalize();
        Path dir =
                root.resolve(OperatorUserPaths.sanitizeOperatorDirName(operator))
                        .toAbsolutePath()
                        .normalize();
        if (!dir.startsWith(root)) {
            return root.resolve(OperatorUserPaths.UNKNOWN_OPERATOR_DIR).normalize();
        }
        return dir;
    }

    /**
     * 比較に使った Excel と加工計画表を操作者フォルダへ保存する。失敗時は empty（例外は投げない）。
     */
    public static Optional<Path> save(
            Map<String, String> ui,
            Path excelPath,
            PlanInputTabularIo.TabularSheet planTab,
            String result,
            String badgeText,
            int diffCount,
            Optional<Path> excelSourcePath,
            Optional<Path> planSourcePath) {
        if (excelPath == null || !Files.isRegularFile(excelPath) || planTab == null) {
            return Optional.empty();
        }
        String operator = OperatorUserPaths.resolveOperatorUser(ui);
        Path operatorDir = resolveOperatorDir(ui, operator);
        try {
            Files.createDirectories(operatorDir);
            Path dest = uniqueSnapshotDir(operatorDir, OffsetDateTime.now(ZoneId.systemDefault()));
            Files.createDirectories(dest);
            Files.copy(
                    excelPath,
                    dest.resolve(EXCEL_FILE),
                    StandardCopyOption.REPLACE_EXISTING);
            JsonTableIo.saveArrayTable(
                    dest.resolve(PLAN_JSON_FILE),
                    planTab.headers() != null ? planTab.headers() : List.of(),
                    planTab.rows() != null ? planTab.rows() : List.of());
            Meta meta =
                    new Meta(
                            OffsetDateTime.now(ZoneId.systemDefault()).toString(),
                            operator,
                            result != null ? result : "",
                            badgeText != null ? badgeText : "",
                            Math.max(0, diffCount),
                            pathString(excelSourcePath.orElse(excelPath)),
                            pathString(planSourcePath.orElse(null)),
                            EXCEL_FILE,
                            PLAN_JSON_FILE);
            Files.writeString(
                    dest.resolve(META_FILE),
                    JSON.writerWithDefaultPrettyPrinter().writeValueAsString(meta),
                    StandardCharsets.UTF_8);
            prune(operatorDir);
            return Optional.of(dest);
        } catch (IOException ex) {
            return Optional.empty();
        }
    }

    public static List<SnapshotRef> listNewestFirst(Map<String, String> ui, String operator) {
        Path opDir = resolveOperatorDir(ui, operator);
        if (!Files.isDirectory(opDir)) {
            return List.of();
        }
        List<Path> dirs = new ArrayList<>();
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(opDir)) {
            for (Path p : stream) {
                if (Files.isDirectory(p) && Files.isRegularFile(p.resolve(META_FILE))) {
                    dirs.add(p);
                }
            }
        } catch (IOException ex) {
            return List.of();
        }
        dirs.sort(Comparator.comparingLong(IdentityCheckHistoryStore::lastModifiedMillis).reversed());
        List<SnapshotRef> out = new ArrayList<>();
        for (Path dir : dirs) {
            Optional<Meta> meta = readMeta(dir);
            meta.ifPresent(m -> out.add(new SnapshotRef(dir, m)));
        }
        return List.copyOf(out);
    }

    public static List<String> listOperatorDirNames(Map<String, String> ui) {
        Path root = resolveRoot(ui);
        if (!Files.isDirectory(root)) {
            return List.of();
        }
        List<String> out = new ArrayList<>();
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(root)) {
            for (Path p : stream) {
                if (Files.isDirectory(p)) {
                    out.add(p.getFileName().toString());
                }
            }
        } catch (IOException ex) {
            return List.of();
        }
        out.sort(String.CASE_INSENSITIVE_ORDER);
        return List.copyOf(out);
    }

    public static Optional<Meta> readMeta(Path snapshotDir) {
        if (snapshotDir == null) {
            return Optional.empty();
        }
        Path meta = snapshotDir.resolve(META_FILE);
        if (!Files.isRegularFile(meta)) {
            return Optional.empty();
        }
        try {
            return Optional.of(JSON.readValue(Files.readString(meta, StandardCharsets.UTF_8), Meta.class));
        } catch (IOException ex) {
            return Optional.empty();
        }
    }

    static Path uniqueSnapshotDir(Path operatorDir, OffsetDateTime when) throws IOException {
        String base = FOLDER_TS.format(when);
        Path candidate = operatorDir.resolve(base);
        if (!Files.exists(candidate)) {
            return candidate;
        }
        for (int i = 2; i < 1000; i++) {
            Path next = operatorDir.resolve(base + "-" + i);
            if (!Files.exists(next)) {
                return next;
            }
        }
        throw new IOException("同一化チェック履歴フォルダ名を確保できません: " + operatorDir);
    }

    static void prune(Path operatorDir) throws IOException {
        if (operatorDir == null || !Files.isDirectory(operatorDir)) {
            return;
        }
        List<Path> dirs = new ArrayList<>();
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(operatorDir)) {
            for (Path p : stream) {
                if (Files.isDirectory(p)) {
                    dirs.add(p);
                }
            }
        }
        dirs.sort(Comparator.comparingLong(IdentityCheckHistoryStore::lastModifiedMillis));
        while (dirs.size() > MAX_SNAPSHOTS_PER_USER) {
            deleteRecursive(dirs.removeFirst());
        }
    }

    private static void deleteRecursive(Path root) throws IOException {
        if (root == null || !Files.exists(root)) {
            return;
        }
        Files.walkFileTree(
                root,
                new SimpleFileVisitor<>() {
                    @Override
                    public FileVisitResult visitFile(Path file, BasicFileAttributes attrs)
                            throws IOException {
                        Files.deleteIfExists(file);
                        return FileVisitResult.CONTINUE;
                    }

                    @Override
                    public FileVisitResult postVisitDirectory(Path dir, IOException exc)
                            throws IOException {
                        Files.deleteIfExists(dir);
                        return FileVisitResult.CONTINUE;
                    }
                });
    }

    private static long lastModifiedMillis(Path p) {
        try {
            return Files.getLastModifiedTime(p).toMillis();
        } catch (IOException e) {
            return Long.MIN_VALUE;
        }
    }

    private static String pathString(Path p) {
        return p != null ? p.toAbsolutePath().normalize().toString() : "";
    }
}
