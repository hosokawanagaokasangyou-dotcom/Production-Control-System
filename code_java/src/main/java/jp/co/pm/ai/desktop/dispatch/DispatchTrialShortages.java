package jp.co.pm.ai.desktop.dispatch;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * {@code dispatch_trial_shortages.json}（配台試行の不足情報・成果パス）の読取り。
 */
public final class DispatchTrialShortages {

    private static final ObjectMapper JSON = new ObjectMapper();

    private DispatchTrialShortages() {}

    /**
     * OP/AS 不足などタスク単位のヒント（{@code reason} および日付・工程の補足）。
     */
    public record ShortageHint(String taskId, String reason, String detail) {

        public String displayLine() {
            String r = reason != null ? reason.trim() : "";
            String d = detail != null ? detail.trim() : "";
            if (r.isEmpty() && d.isEmpty()) {
                return "";
            }
            if (d.isEmpty()) {
                return r;
            }
            if (r.isEmpty()) {
                return d;
            }
            return r + "（" + d + "）";
        }
    }

    public record Paths(String productionPlan, String memberSchedule) {}

    /** {@link #read(Path)} に加え、{@code op_shortage} / {@code as_shortage} を解析したもの。 */
    public record FullBundle(Paths paths, List<ShortageHint> shortageHints) {}

    /**
     * @param shortageJsonPath {@code dispatch_interactive_trial.py} が書く JSON
     */
    public static Paths read(Path shortageJsonPath) throws IOException {
        if (shortageJsonPath == null || !Files.isRegularFile(shortageJsonPath)) {
            return new Paths("", "");
        }
        JsonNode root = readRoot(shortageJsonPath);
        return pathsFromRoot(root);
    }

    public static FullBundle readFull(Path shortageJsonPath) throws IOException {
        if (shortageJsonPath == null || !Files.isRegularFile(shortageJsonPath)) {
            return new FullBundle(new Paths("", ""), List.of());
        }
        JsonNode root = readRoot(shortageJsonPath);
        Paths paths = pathsFromRoot(root);
        List<ShortageHint> hints = new ArrayList<>();
        hints.addAll(parseShortageArray(root, "op_shortage"));
        hints.addAll(parseShortageArray(root, "as_shortage"));
        return new FullBundle(paths, List.copyOf(hints));
    }

    private static JsonNode readRoot(Path shortageJsonPath) throws IOException {
        return JSON.readTree(Files.readString(shortageJsonPath, StandardCharsets.UTF_8));
    }

    private static Paths pathsFromRoot(JsonNode root) {
        String plan = text(root, "production_plan");
        String member = text(root, "member_schedule");
        return new Paths(plan != null ? plan : "", member != null ? member : "");
    }

    private static List<ShortageHint> parseShortageArray(JsonNode root, String field) {
        JsonNode arr = root != null ? root.get(field) : null;
        if (arr == null || !arr.isArray()) {
            return List.of();
        }
        List<ShortageHint> out = new ArrayList<>();
        for (JsonNode n : arr) {
            if (n == null || !n.isObject()) {
                continue;
            }
            String tid = text(n, "task_id");
            String reason = text(n, "reason");
            String date = text(n, "date");
            String proc = text(n, "process");
            String machineName = text(n, "machine_name");
            String detail =
                    Stream.of(date, proc, machineName)
                            .map(s -> s != null ? s.trim() : "")
                            .filter(s -> !s.isEmpty())
                            .collect(Collectors.joining(" "));
            out.add(new ShortageHint(tid, reason, detail));
        }
        return out;
    }

    /** 同一タスクに複数ヒントがあるときは重複を除いて結合する。 */
    public static String mergeHintsForTask(List<ShortageHint> hints, String taskId) {
        if (hints == null || hints.isEmpty()) {
            return "";
        }
        String tid = taskId != null ? taskId.trim() : "";
        if (tid.isEmpty()) {
            return "";
        }
        LinkedHashSet<String> parts = new LinkedHashSet<>();
        for (ShortageHint h : hints) {
            if (!tid.equals(h.taskId() != null ? h.taskId().trim() : "")) {
                continue;
            }
            String line = h.displayLine();
            if (!line.isBlank()) {
                parts.add(line);
            }
        }
        return String.join("；", parts);
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
}
