package jp.co.pm.ai.desktop.dispatch;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Optional;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * {@code 結果_配台表.json} の生成者・生成日時。
 *
 * <p>JSON ルートの {@code generated_by} / {@code generated_at} を優先し、
 * 無い場合はファイルの最終更新時刻を日時の代替とする（旧ファイル互換）。
 */
public final class ResultDispatchProvenance {

    public static final String JSON_KEY_GENERATED_BY = "generated_by";
    public static final String JSON_KEY_GENERATED_AT = "generated_at";

    private static final ObjectMapper JSON = new ObjectMapper();
    private static final DateTimeFormatter DISPLAY_FMT =
            DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm");
    private static final ZoneId ZONE = ZoneId.systemDefault();

    private ResultDispatchProvenance() {}

    /**
     * @param generatedBy 生成操作者（空可）
     * @param generatedAt 生成日時（空可）
     * @param generatedAtFromFileMtime {@code generatedAt} がファイル mtime 由来か
     * @param path 読込元パス（表示用。null 可）
     */
    public record Info(
            String generatedBy,
            Optional<Instant> generatedAt,
            boolean generatedAtFromFileMtime,
            Path path) {

        public Info {
            generatedBy = generatedBy == null ? "" : generatedBy.strip();
            generatedAt = generatedAt != null ? generatedAt : Optional.empty();
        }

        public boolean hasGenerator() {
            return !generatedBy.isBlank()
                    && !"unknown".equalsIgnoreCase(generatedBy)
                    && !"（未記録）".equals(generatedBy);
        }

        /** UI 1 行用。例: {@code 生成: 山田 / 2026/09/11 10:30} */
        public String formatDisplayLine() {
            String who = hasGenerator() ? generatedBy : "（未記録）";
            String when;
            if (generatedAt.isPresent()) {
                String ts = LocalDateTime.ofInstant(generatedAt.get(), ZONE).format(DISPLAY_FMT);
                when = generatedAtFromFileMtime ? "ファイル更新 " + ts : ts;
            } else {
                when = "（不明）";
            }
            return "配台JSON 生成: " + who + " ／ " + when;
        }
    }

    /** パスが無い・読めないときは empty。 */
    public static Optional<Info> read(Path path) {
        if (path == null || !Files.isRegularFile(path)) {
            return Optional.empty();
        }
        String by = "";
        Instant at = null;
        boolean fromMtime = false;
        try {
            String raw = Files.readString(path, StandardCharsets.UTF_8);
            JsonNode root = JSON.readTree(raw);
            by = text(root, JSON_KEY_GENERATED_BY);
            at = parseInstant(text(root, JSON_KEY_GENERATED_AT)).orElse(null);
        } catch (Exception ignored) {
            // 表データは別経路で読む。メタだけ失敗しても mtime に落とす
        }
        if (at == null) {
            try {
                at = Files.getLastModifiedTime(path).toInstant();
                fromMtime = true;
            } catch (Exception ignored) {
                // leave null
            }
        }
        return Optional.of(new Info(by, Optional.ofNullable(at), fromMtime, path));
    }

    public static Optional<Instant> parseInstant(String raw) {
        if (raw == null || raw.isBlank()) {
            return Optional.empty();
        }
        String s = raw.strip();
        try {
            return Optional.of(Instant.parse(s));
        } catch (DateTimeParseException ignored) {
            // continue
        }
        try {
            return Optional.of(LocalDateTime.parse(s).atZone(ZONE).toInstant());
        } catch (DateTimeParseException ignored) {
            // continue
        }
        try {
            return Optional.of(LocalDateTime.parse(s, DISPLAY_FMT).atZone(ZONE).toInstant());
        } catch (DateTimeParseException ignored) {
            return Optional.empty();
        }
    }

    public static String formatInstantIso(Instant instant) {
        if (instant == null) {
            return "";
        }
        return instant.toString();
    }

    private static String text(JsonNode root, String key) {
        if (root == null || key == null) {
            return "";
        }
        JsonNode n = root.get(key);
        if (n == null || n.isNull()) {
            return "";
        }
        return n.asText("").strip();
    }
}
