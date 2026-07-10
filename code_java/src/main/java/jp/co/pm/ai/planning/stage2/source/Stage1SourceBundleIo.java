package jp.co.pm.ai.planning.stage2.source;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Map;
import java.util.Optional;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.fasterxml.jackson.databind.node.ObjectNode;

import jp.co.pm.ai.desktop.config.AppPaths;

/** {@code .pm-ai-cache/stage1-source-bundle.json} の読み書き。 */
public final class Stage1SourceBundleIo {

    public static final String CACHE_FILE_NAME = "stage1-source-bundle.json";

    private static final ObjectMapper JSON =
            new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);

    private static final DateTimeFormatter ISO_LOCAL =
            DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss");

    private Stage1SourceBundleIo() {}

    public static Path defaultCachePath(Map<String, String> ui) {
        return AppPaths.resolveRepoRoot(ui != null ? ui : Map.of())
                .resolve(".pm-ai-cache")
                .resolve(CACHE_FILE_NAME);
    }

    public static void write(Path path, Stage1SourceBundle bundle) throws IOException {
        if (path == null || bundle == null) {
            throw new IOException("bundle または path が null です");
        }
        Files.createDirectories(path.getParent());
        ObjectNode root = JSON.createObjectNode();
        root.put("planExtractionTime", format(bundle.planExtractionTime()));
        root.put("dailyReportExtractionTime", format(bundle.dailyReportExtractionTime()));
        root.put("pairDeltaMinutes", bundle.pairDeltaMinutes());
        root.put("processingPlanPath", bundle.processingPlanPath());
        root.put("dailyReportCsvPath", bundle.dailyReportCsvPath());
        root.put("dataExtractionWorkbookPath", bundle.dataExtractionWorkbookPath());
        root.put("stage1CompletedAtEpochMillis", bundle.stage1CompletedAtEpochMillis());
        JSON.writeValue(path.toFile(), root);
    }

    public static void writeDefault(Map<String, String> ui, Stage1SourceBundle bundle)
            throws IOException {
        write(defaultCachePath(ui), bundle);
    }

    public static Optional<Stage1SourceBundle> readIfPresent(Map<String, String> ui) {
        Path path = defaultCachePath(ui);
        if (!Files.isRegularFile(path)) {
            return Optional.empty();
        }
        try {
            return Optional.of(read(path));
        } catch (IOException ex) {
            return Optional.empty();
        }
    }

    public static Stage1SourceBundle read(Path path) throws IOException {
        ObjectNode root = (ObjectNode) JSON.readTree(path.toFile());
        return new Stage1SourceBundle(
                parseDateTime(text(root, "planExtractionTime")),
                parseDateTime(text(root, "dailyReportExtractionTime")),
                root.path("pairDeltaMinutes").asLong(0L),
                text(root, "processingPlanPath"),
                text(root, "dailyReportCsvPath"),
                text(root, "dataExtractionWorkbookPath"),
                root.path("stage1CompletedAtEpochMillis").asLong(0L));
    }

    public static void deleteIfExists(Map<String, String> ui) {
        try {
            Files.deleteIfExists(defaultCachePath(ui));
        } catch (IOException ignored) {
            // best effort
        }
    }

    private static String text(ObjectNode node, String field) {
        if (node == null || field == null) {
            return "";
        }
        var n = node.get(field);
        return n != null && !n.isNull() ? n.asText("").strip() : "";
    }

    private static String format(LocalDateTime dt) {
        return dt != null ? dt.format(ISO_LOCAL) : "";
    }

    private static LocalDateTime parseDateTime(String raw) {
        if (raw == null || raw.isBlank()) {
            return null;
        }
        return LocalDateTime.parse(raw.strip(), ISO_LOCAL);
    }
}
