package jp.co.pm.ai.planning.stage2.source;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.StandardCopyOption;
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

    @FunctionalInterface
    interface MoveOperation {
        void move(Path temporary, Path target) throws IOException;
    }

    public static Path defaultCachePath(Map<String, String> ui) {
        return AppPaths.resolveRepoRoot(ui != null ? ui : Map.of())
                .resolve(".pm-ai-cache")
                .resolve(CACHE_FILE_NAME);
    }

    public static void write(Path path, Stage1SourceBundle bundle) throws IOException {
        writeWithMove(path, bundle, Stage1SourceBundleIo::replaceAtomically);
    }

    static void writeWithMove(Path path, Stage1SourceBundle bundle, MoveOperation mover) throws IOException {
        if (path == null || bundle == null) {
            throw new IOException("bundle または path が null です");
        }
        var error = bundle.validationError();
        if (error.isPresent()) throw new IOException(error.get());
        Path target = path.toAbsolutePath().normalize();
        Files.createDirectories(target.getParent());
        Path temporary = Files.createTempFile(target.getParent(), "." + target.getFileName() + ".", ".tmp");
        ObjectNode root = JSON.createObjectNode();
        root.put("planExtractionTime", format(bundle.planExtractionTime()));
        root.put("dailyReportExtractionTime", format(bundle.dailyReportExtractionTime()));
        root.put("pairDeltaMinutes", bundle.pairDeltaMinutes());
        root.put("processingPlanPath", bundle.processingPlanPath());
        root.put("dailyReportCsvPath", bundle.dailyReportCsvPath());
        root.put("dataExtractionWorkbookPath", bundle.dataExtractionWorkbookPath());
        root.put("stage1CompletedAtEpochMillis", bundle.stage1CompletedAtEpochMillis());
        try {
            JSON.writeValue(temporary.toFile(), root);
            mover.move(temporary, target);
        } catch (IOException | RuntimeException ex) {
            try { Files.deleteIfExists(target); } catch (IOException cleanup) { ex.addSuppressed(cleanup); }
            throw ex;
        } finally {
            Files.deleteIfExists(temporary);
        }
    }

    private static void replaceAtomically(Path temporary, Path target) throws IOException {
        try {
            Files.move(temporary, target, StandardCopyOption.ATOMIC_MOVE, StandardCopyOption.REPLACE_EXISTING);
        } catch (AtomicMoveNotSupportedException ex) {
            Files.move(temporary, target, StandardCopyOption.REPLACE_EXISTING);
        }
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

    public static Optional<Stage1SourceBundle> readIfPresentStrict(Map<String, String> ui)
            throws IOException {
        Path path = defaultCachePath(ui);
        return Files.isRegularFile(path) ? Optional.of(read(path)) : Optional.empty();
    }

    public static Stage1SourceBundle read(Path path) throws IOException {
        try {
            var node = JSON.readTree(path.toFile());
            if (!(node instanceof ObjectNode root)) throw new IOException("bundle JSON がobjectではありません");
            Stage1SourceBundle bundle = new Stage1SourceBundle(
                    parseRequiredDateTime(root, "planExtractionTime"),
                    parseRequiredDateTime(root, "dailyReportExtractionTime"),
                    requiredLong(root, "pairDeltaMinutes"),
                    requiredText(root, "processingPlanPath"),
                    requiredText(root, "dailyReportCsvPath"),
                    requiredText(root, "dataExtractionWorkbookPath"),
                    requiredLong(root, "stage1CompletedAtEpochMillis"));
            var error = bundle.validationError();
            if (error.isPresent()) throw new IOException(error.get());
            return bundle;
        } catch (IOException ex) {
            throw ex;
        } catch (RuntimeException ex) {
            throw new IOException("bundle JSON が不正です: " + ex.getMessage(), ex);
        }
    }

    public static void deleteIfExists(Map<String, String> ui) {
        try {
            Files.deleteIfExists(defaultCachePath(ui));
        } catch (IOException ignored) {
            // best effort
        }
    }

    public static void deleteDefaultIfExists(Map<String, String> ui) throws IOException {
        Files.deleteIfExists(defaultCachePath(ui));
    }

    private static String text(ObjectNode node, String field) {
        if (node == null || field == null) {
            return "";
        }
        var n = node.get(field);
        return n != null && !n.isNull() ? n.asText("").strip() : "";
    }

    private static String requiredText(ObjectNode node, String field) throws IOException {
        var valueNode = node.get(field);
        if (valueNode == null || !valueNode.isTextual()) {
            throw new IOException(field + " は文字列である必要があります");
        }
        String value = valueNode.textValue().strip();
        if (value.isBlank()) throw new IOException(field + " がありません");
        return value;
    }

    private static long requiredLong(ObjectNode node, String field) throws IOException {
        var value = node.get(field);
        if (value == null || !value.isIntegralNumber() || !value.canConvertToLong()) {
            throw new IOException(field + " は整数である必要があります");
        }
        return value.longValue();
    }

    private static LocalDateTime parseRequiredDateTime(ObjectNode node, String field) throws IOException {
        try {
            return LocalDateTime.parse(requiredText(node, field), ISO_LOCAL);
        } catch (RuntimeException ex) {
            throw new IOException(field + " が不正です", ex);
        }
    }

    private static String format(LocalDateTime dt) {
        return dt != null ? dt.format(ISO_LOCAL) : "";
    }

}
