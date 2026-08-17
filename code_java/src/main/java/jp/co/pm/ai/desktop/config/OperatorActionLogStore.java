package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.time.Duration;
import java.time.Instant;
import java.time.LocalDate;
import java.time.OffsetDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Map;

import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * 工場共有 DATA 配下の操作者別操作ログ（日次 NDJSON、保持90日）。
 */
public final class OperatorActionLogStore {

    public static final int RETENTION_DAYS = 90;

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final DateTimeFormatter DAY = DateTimeFormatter.ISO_LOCAL_DATE;

    private OperatorActionLogStore() {}

    public record Entry(String ts, String operator, String action, String result, String detail) {}

    public static Path resolveRoot(Map<String, String> ui) {
        return AppPaths.resolveOperatorActionLogRoot(ui);
    }

    public static Path resolveOperatorDir(Map<String, String> ui, String operator) {
        return resolveRoot(ui).resolve(OperatorUserPaths.sanitizeOperatorDirName(operator));
    }

    public static Path resolveDailyFile(Map<String, String> ui, String operator, LocalDate day) {
        LocalDate d = day != null ? day : LocalDate.now(ZoneId.systemDefault());
        return resolveOperatorDir(ui, operator).resolve(DAY.format(d) + ".ndjson");
    }

    public static boolean append(
            Map<String, String> ui, String operator, String action, String result, String detail) {
        String name = operator != null ? operator.strip() : "";
        if (name.isEmpty()) {
            name = OperatorUserPaths.UNKNOWN_OPERATOR_DIR;
        }
        Path file = resolveDailyFile(ui, name, LocalDate.now(ZoneId.systemDefault()));
        Entry entry =
                new Entry(
                        OffsetDateTime.now().toString(),
                        name,
                        action != null ? action : "",
                        result != null ? result : "",
                        detail != null ? detail : "");
        try {
            Files.createDirectories(file.getParent());
            String line = JSON.writeValueAsString(entry) + "\n";
            Files.writeString(
                    file,
                    line,
                    StandardCharsets.UTF_8,
                    StandardOpenOption.CREATE,
                    StandardOpenOption.APPEND);
            return true;
        } catch (IOException ex) {
            return false;
        }
    }

    public static List<String> listOperators(Map<String, String> ui) {
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
        return out;
    }

    public static List<Entry> readOperator(Map<String, String> ui, String operator, Instant now) {
        Path dir = resolveOperatorDir(ui, operator);
        pruneOlderThan(dir, now != null ? now : Instant.now(), RETENTION_DAYS);
        if (!Files.isDirectory(dir)) {
            return List.of();
        }
        List<Path> files = new ArrayList<>();
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(dir, "*.ndjson")) {
            for (Path p : stream) {
                if (Files.isRegularFile(p)) {
                    files.add(p);
                }
            }
        } catch (IOException ex) {
            return List.of();
        }
        files.sort(Comparator.comparing((Path p) -> p.getFileName().toString()).reversed());
        List<Entry> out = new ArrayList<>();
        for (Path file : files) {
            out.addAll(readFileNewestFirst(file));
        }
        return out;
    }

    public static int pruneOlderThan(Path operatorDir, Instant now, int days) {
        if (operatorDir == null || !Files.isDirectory(operatorDir) || now == null || days < 0) {
            return 0;
        }
        Instant cutoff = now.minus(Duration.ofDays(days));
        int removed = 0;
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(operatorDir, "*.ndjson")) {
            for (Path p : stream) {
                if (!Files.isRegularFile(p)) {
                    continue;
                }
                Instant mtime;
                try {
                    mtime = Files.getLastModifiedTime(p).toInstant();
                } catch (IOException ex) {
                    continue;
                }
                if (mtime.isBefore(cutoff)) {
                    try {
                        Files.deleteIfExists(p);
                        removed++;
                    } catch (IOException ignored) {
                        // 個別削除失敗はスキップ
                    }
                }
            }
        } catch (IOException ex) {
            return removed;
        }
        return removed;
    }

    private static List<Entry> readFileNewestFirst(Path file) {
        List<Entry> rows = new ArrayList<>();
        try {
            for (String line : Files.readAllLines(file, StandardCharsets.UTF_8)) {
                if (line == null || line.isBlank()) {
                    continue;
                }
                try {
                    rows.add(JSON.readValue(line, Entry.class));
                } catch (IOException ignored) {
                    // 壊れた行はスキップ
                }
            }
        } catch (IOException ex) {
            return List.of();
        }
        List<Entry> newestFirst = new ArrayList<>(rows.size());
        for (int i = rows.size() - 1; i >= 0; i--) {
            newestFirst.add(rows.get(i));
        }
        return newestFirst;
    }
}
