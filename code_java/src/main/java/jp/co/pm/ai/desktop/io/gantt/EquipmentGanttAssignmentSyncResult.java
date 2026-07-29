package jp.co.pm.ai.desktop.io.gantt;

import java.util.ArrayList;
import java.util.List;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

/** {@code sync_equipment_gantt_assignment.py} の 1 行 JSON 応答。 */
public final class EquipmentGanttAssignmentSyncResult {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    private final boolean ok;
    private final String status;
    private final String timelineHash;
    private final String confirmToken;
    private final String detail;
    private final String contractPath;
    private final String planXlsxPath;
    private final List<Issue> errors;
    private final List<Issue> warnings;
    private final List<String> backupPaths;

    private EquipmentGanttAssignmentSyncResult(
            boolean ok,
            String status,
            String timelineHash,
            String confirmToken,
            String detail,
            String contractPath,
            String planXlsxPath,
            List<Issue> errors,
            List<Issue> warnings,
            List<String> backupPaths) {
        this.ok = ok;
        this.status = status != null ? status : "";
        this.timelineHash = timelineHash != null ? timelineHash : "";
        this.confirmToken = confirmToken != null ? confirmToken : "";
        this.detail = detail != null ? detail : "";
        this.contractPath = contractPath != null ? contractPath : "";
        this.planXlsxPath = planXlsxPath != null ? planXlsxPath : "";
        this.errors = List.copyOf(errors != null ? errors : List.of());
        this.warnings = List.copyOf(warnings != null ? warnings : List.of());
        this.backupPaths = List.copyOf(backupPaths != null ? backupPaths : List.of());
    }

    public static EquipmentGanttAssignmentSyncResult parseJson(String raw) throws Exception {
        if (raw == null || raw.isBlank()) {
            throw new IllegalArgumentException("empty sync response");
        }
        String line = raw.strip();
        int nl = line.lastIndexOf('\n');
        if (nl >= 0) {
            line = line.substring(nl + 1).strip();
        }
        JsonNode root = MAPPER.readTree(line);
        List<Issue> errors = parseIssues(root.get("errors"));
        List<Issue> warnings = parseIssues(root.get("warnings"));
        List<String> backups = new ArrayList<>();
        JsonNode bp = root.get("backup_paths");
        if (bp != null && bp.isArray()) {
            for (JsonNode n : bp) {
                if (n != null && n.isTextual()) {
                    backups.add(n.asText());
                }
            }
        }
        return new EquipmentGanttAssignmentSyncResult(
                root.path("ok").asBoolean(false),
                root.path("status").asText(""),
                root.path("timeline_hash").asText(""),
                root.path("confirm_token").asText(""),
                root.path("detail").asText(root.path("error").asText("")),
                root.path("contract_path").asText(""),
                root.path("plan_xlsx_path").asText(""),
                errors,
                warnings,
                backups);
    }

    private static List<Issue> parseIssues(JsonNode arr) {
        if (arr == null || !arr.isArray()) {
            return List.of();
        }
        List<Issue> out = new ArrayList<>();
        for (JsonNode n : arr) {
            if (n == null || !n.isObject()) {
                continue;
            }
            out.add(
                    new Issue(
                            n.path("code").asText(""),
                            n.path("message").asText(""),
                            n.path("person").asText("")));
        }
        return out;
    }

    public boolean ok() {
        return ok;
    }

    public String status() {
        return status;
    }

    public boolean hasWarnings() {
        return "warnings".equals(status) || !warnings.isEmpty();
    }

    public String timelineHash() {
        return timelineHash;
    }

    public String confirmToken() {
        return confirmToken;
    }

    public String detail() {
        return detail;
    }

    public String contractPath() {
        return contractPath;
    }

    public String planXlsxPath() {
        return planXlsxPath;
    }

    public List<Issue> errors() {
        return errors;
    }

    public List<Issue> warnings() {
        return warnings;
    }

    public List<String> backupPaths() {
        return backupPaths;
    }

    public String formatIssuesForDialog() {
        StringBuilder sb = new StringBuilder();
        for (Issue w : warnings) {
            if (!w.message().isBlank()) {
                sb.append("・").append(w.message()).append('\n');
            }
        }
        for (Issue e : errors) {
            if (!e.message().isBlank()) {
                sb.append("・").append(e.message()).append('\n');
            }
        }
        if (!detail.isBlank()) {
            sb.append(detail);
        }
        return sb.toString().strip();
    }

    public record Issue(String code, String message, String person) {}
}
