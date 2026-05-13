package jp.co.pm.ai.planning.stage2.input;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Locale;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

/** 配台不要ルール JSON の件数把握（Python 経路のログ突合用）。 */
public final class Stage2ExcludeRulesReader {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    private Stage2ExcludeRulesReader() {}

    public static int countRules(Path path) throws IOException {
        if (path == null || !Files.isRegularFile(path)) {
            return 0;
        }
        JsonNode root = MAPPER.readTree(Files.readString(path, StandardCharsets.UTF_8));
        if (root == null || root.isNull()) {
            return 0;
        }
        if (root.isArray()) {
            return root.size();
        }
        JsonNode rules = root.get("rules");
        if (rules != null && rules.isArray()) {
            return rules.size();
        }
        return 0;
    }

    public static String fileLabel(Path path) {
        if (path == null) {
            return "(なし)";
        }
        return path.getFileName() != null
                ? path.getFileName().toString().toLowerCase(Locale.ROOT)
                : path.toString();
    }
}
