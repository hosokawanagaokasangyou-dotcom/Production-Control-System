package jp.co.pm.ai.desktop.dispatch;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * {@code dispatch_trial_shortages.json}（配台試行の不足情報・成果パス）の最小読取り。
 */
public final class DispatchTrialShortages {

    private static final ObjectMapper JSON = new ObjectMapper();

    private DispatchTrialShortages() {}

    /**
     * @param shortageJsonPath {@code dispatch_interactive_trial.py} が書く JSON
     */
    public static Paths read(Path shortageJsonPath) throws IOException {
        if (shortageJsonPath == null || !Files.isRegularFile(shortageJsonPath)) {
            return new Paths("", "");
        }
        JsonNode root = JSON.readTree(Files.readString(shortageJsonPath, StandardCharsets.UTF_8));
        String plan = text(root, "production_plan");
        String member = text(root, "member_schedule");
        return new Paths(plan != null ? plan : "", member != null ? member : "");
    }

    private static String text(JsonNode root, String field) {
        if (root == null || field == null) {
            return "";
        }
        JsonNode n = root.get(field);
        if (n == null || n.isNull()) {
            return "";
        }
        return n.asText("").trim();
    }

    public record Paths(String productionPlan, String memberSchedule) {}
}
