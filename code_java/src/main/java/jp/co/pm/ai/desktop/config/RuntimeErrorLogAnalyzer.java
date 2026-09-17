package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * 現工場 {@code remote_log/<操作者>/} の日次 UI ログと世代 meta から実行時エラー行を集める。
 */
public final class RuntimeErrorLogAnalyzer {

    private static final ObjectMapper JSON = new ObjectMapper();

    private RuntimeErrorLogAnalyzer() {}

    public record Row(
            String ts,
            String operator,
            String factory,
            String appVersion,
            String osUser,
            String host,
            String eventId,
            String severity,
            String excerpt,
            String sourcePath) {}

    public static String severityOf(String line) {
        if (line == null) {
            return "warn";
        }
        if (line.contains("Exception")
                || line.contains("失敗")
                || line.contains("読取不可")
                || line.contains("ドロップされたファイルがありません")) {
            return "error";
        }
        return "warn";
    }

    public static List<String> listOperators(Map<String, String> ui) {
        Path root = AppPaths.resolveRemoteLogRoot(ui);
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

    public static List<Row> scan(Map<String, String> ui) {
        List<Row> rows = new ArrayList<>();
        Path root = AppPaths.resolveRemoteLogRoot(ui);
        if (!Files.isDirectory(root)) {
            return rows;
        }
        List<String> operators = listOperators(ui);
        for (String operator : operators) {
            Path userDir = root.resolve(operator);
            collectDaily(userDir, operator, rows);
            collectGenerations(userDir, operator, rows);
        }
        rows.sort(Comparator.comparing(Row::ts).reversed());
        return rows;
    }

    private static void collectDaily(Path userDir, String operator, List<Row> rows) {
        Path dailyDir = userDir.resolve(RemoteSupportLogArchive.UI_DAILY_DIR_NAME);
        if (!Files.isDirectory(dailyDir)) {
            return;
        }
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(dailyDir, "*.txt")) {
            for (Path file : stream) {
                collectUiFile(file, operator, "ui_daily", rows);
            }
        } catch (IOException ignored) {
            // 個別ファイルはスキップ
        }
    }

    private static void collectGenerations(Path userDir, String operator, List<Row> rows) {
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(userDir)) {
            for (Path child : stream) {
                if (!Files.isDirectory(child)
                        || RemoteSupportLogArchive.UI_DAILY_DIR_NAME.equals(
                                child.getFileName().toString())) {
                    continue;
                }
                collectMeta(child, operator, rows);
                Path uiLog = child.resolve(RemoteSupportLogArchive.UI_RUN_LOG_FILENAME);
                if (Files.isRegularFile(uiLog)) {
                    collectUiFile(uiLog, operator, generationEventId(child), rows);
                }
            }
        } catch (IOException ignored) {
            // 世代フォルダ列挙失敗はスキップ
        }
    }

    private static String generationEventId(Path genDir) {
        String name = genDir.getFileName().toString();
        int us = name.indexOf('_');
        if (us < 0 || us + 1 >= name.length()) {
            return "runtime";
        }
        return name.substring(us + 1);
    }

    private static void collectMeta(Path genDir, String operator, List<Row> rows) {
        Path meta = genDir.resolve(RemoteSupportLogArchive.META_JSON_FILENAME);
        if (!Files.isRegularFile(meta)) {
            return;
        }
        try {
            JsonNode root = JSON.readTree(meta.toFile());
            String error = text(root, "error");
            if (error.isBlank()) {
                return;
            }
            rows.add(
                    new Row(
                            text(root, "saved_at"),
                            operator,
                            text(root, "factory"),
                            text(root, "app_version"),
                            text(root, "os_user"),
                            text(root, "host"),
                            text(root, "stage_id"),
                            "error",
                            error,
                            meta.toAbsolutePath().normalize().toString()));
        } catch (IOException ignored) {
            // 壊れた meta はスキップ
        }
    }

    private static void collectUiFile(
            Path file, String operator, String defaultEventId, List<Row> rows) {
        List<String> lines;
        try {
            lines = Files.readAllLines(file, StandardCharsets.UTF_8);
        } catch (IOException ex) {
            return;
        }
        Header h = Header.parse(lines, operator);
        String source = file.toAbsolutePath().normalize().toString();
        Set<String> seen = new LinkedHashSet<>();
        for (String line : lines) {
            if (line == null || line.startsWith("#")) {
                continue;
            }
            if (!RemoteSupportLogArchive.isDiagnosticRuntimeLine(line)) {
                continue;
            }
            String excerpt = line.strip();
            if (excerpt.isEmpty() || !seen.add(excerpt)) {
                continue;
            }
            rows.add(
                    new Row(
                            h.ts,
                            h.operator,
                            h.factory,
                            h.appVersion,
                            h.osUser,
                            h.host,
                            eventIdOf(line, defaultEventId),
                            severityOf(line),
                            excerpt,
                            source));
        }
    }

    private static String eventIdOf(String line, String fallback) {
        String id = RemoteSupportLogArchive.eventIdForDiagnosticLine(line);
        if (RemoteSupportLogArchive.STAGE_RUNTIME.equals(id)
                && fallback != null
                && !fallback.isBlank()) {
            return fallback;
        }
        return id;
    }

    private static String text(JsonNode root, String field) {
        if (root == null || !root.has(field) || root.get(field).isNull()) {
            return "";
        }
        String v = root.get(field).asText("");
        return v != null ? v : "";
    }

    private static final class Header {
        private final String ts;
        private final String operator;
        private final String factory;
        private final String appVersion;
        private final String osUser;
        private final String host;

        private Header(
                String ts,
                String operator,
                String factory,
                String appVersion,
                String osUser,
                String host) {
            this.ts = ts;
            this.operator = operator;
            this.factory = factory;
            this.appVersion = appVersion;
            this.osUser = osUser;
            this.host = host;
        }

        static Header parse(List<String> lines, String operatorFallback) {
            String factory = "";
            String appVersion = "";
            String osUser = "";
            String host = "";
            String ts = "";
            String operator = operatorFallback;
            for (String line : lines) {
                if (line == null || !line.startsWith("#")) {
                    continue;
                }
                if (line.startsWith("# pm-ai-desktop ui_daily ")) {
                    ts = line.substring("# pm-ai-desktop ui_daily ".length()).strip();
                }
                for (String part : line.split("\\s+")) {
                    int eq = part.indexOf('=');
                    if (eq <= 0) {
                        continue;
                    }
                    String key = part.substring(0, eq);
                    String val = part.substring(eq + 1);
                    switch (key) {
                        case "operator" -> operator = val;
                        case "factory" -> factory = val;
                        case "app_version" -> appVersion = val;
                        case "os_user" -> osUser = val;
                        case "host" -> host = val;
                        default -> {}
                    }
                }
            }
            if (ts.isBlank()) {
                ts = "";
            }
            return new Header(ts, operator, factory, appVersion, osUser, host);
        }
    }
}
