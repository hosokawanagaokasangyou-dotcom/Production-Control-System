package jp.co.pm.ai.desktop.dispatch;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * 段階2.1 試行直前の段階2配台数量（プロファイル×配台日）と勤怠上書きメタ。
 * 配台計画手動修正タブの (段階2後)/(段階2.1後) 比較・勤怠適用バナー用。
 */
public final class Stage21TrialSnapshotStore {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final DateTimeFormatter APPLIED_AT_FMT = DateTimeFormatter.ISO_LOCAL_DATE_TIME;

    private static final String SIDECAR_SUFFIX = ".stage21_trial_meta.json";

    private Stage21TrialSnapshotStore() {}

    public record OverrideSummary(int workOn, int workOff, int overtimeCells) {

        public static OverrideSummary empty() {
            return new OverrideSummary(0, 0, 0);
        }

        public int totalChanges() {
            return workOn + workOff + overtimeCells;
        }

        public String formatSummaryLine() {
            return "休日出勤（○化）: "
                    + workOn
                    + " セル / 休日扱い（グレー化）: "
                    + workOff
                    + " セル / 残業時間の変更: "
                    + overtimeCells
                    + " セル";
        }
    }

    public record Stage21TrialMeta(
            boolean stage21Applied,
            String stage21ResultDispatchJson,
            String overtimeOverridesJson,
            OverrideSummary overrideSummary,
            String appliedAt,
            Map<String, Double> entries) {

        public static Stage21TrialMeta empty() {
            return new Stage21TrialMeta(false, "", "", OverrideSummary.empty(), "", Map.of());
        }

        public boolean hasTrialApplied() {
            return stage21Applied || (entries != null && !entries.isEmpty());
        }

        public Path overtimeOverridesPath() {
            if (overtimeOverridesJson == null || overtimeOverridesJson.isBlank()) {
                return null;
            }
            return Path.of(overtimeOverridesJson);
        }

        public Path stage21ResultDispatchPath() {
            if (stage21ResultDispatchJson == null || stage21ResultDispatchJson.isBlank()) {
                return null;
            }
            return Path.of(stage21ResultDispatchJson);
        }
    }

    public static Path sidecarPathFor(Path mainResultDispatchJson) {
        if (mainResultDispatchJson == null) {
            return null;
        }
        String fileName = mainResultDispatchJson.getFileName().toString();
        if (!fileName.endsWith(".json")) {
            return mainResultDispatchJson.resolveSibling(fileName + SIDECAR_SUFFIX);
        }
        return mainResultDispatchJson.resolveSibling(
                fileName.substring(0, fileName.length() - 5) + SIDECAR_SUFFIX);
    }

    /** 段階2 正本表の当日配台数量（暦日×m）を baseline として記録する。 */
    public static Map<String, Double> captureStage2PlanFromDocument(
            ResultDispatchDocument sourceDoc, List<LocalDate> axis) {
        Map<String, Double> out = new LinkedHashMap<>();
        if (sourceDoc == null || axis == null || axis.isEmpty()) {
            return out;
        }
        List<Map<String, String>> profiles =
                ResultDispatchPivot.distinctWideTaskProfiles(
                        sourceDoc.columns(),
                        sourceDoc.rows(),
                        ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
        for (Map<String, String> profile : profiles) {
            for (LocalDate day : axis) {
                double plan =
                        ResultDispatchPivot.sumQuantityForProfileAndDateForWideMerge(
                                sourceDoc.rows(),
                                profile,
                                day,
                                ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
                String key = cellKey(profile, day);
                if (plan > 1e-3) {
                    out.put(key, plan);
                }
            }
        }
        return out;
    }

    public static String cellKey(Map<String, String> profile, LocalDate day) {
        if (profile == null || day == null) {
            return "";
        }
        return DispatchTrialShortages.wideShortfallKey(
                profile.getOrDefault("依頼NO", ""),
                profile.get(ResultDispatchSchema.COL_MACHINE),
                day.toString());
    }

    public static void writeWithMeta(
            Path mainResultDispatchJson,
            Map<String, Double> snapshot,
            Path stage21ResultDispatchJson,
            Path overtimeOverridesJson,
            OverrideSummary overrideSummary) {
        Path sidecar = sidecarPathFor(mainResultDispatchJson);
        if (sidecar == null) {
            return;
        }
        try {
            ObjectNode root = JSON.createObjectNode();
            root.put(
                    "source_json",
                    mainResultDispatchJson.toAbsolutePath().normalize().toString());
            root.put("stage21_applied", true);
            if (stage21ResultDispatchJson != null) {
                root.put(
                        "stage21_result_dispatch_json",
                        stage21ResultDispatchJson.toAbsolutePath().normalize().toString());
            }
            if (overtimeOverridesJson != null) {
                root.put(
                        "overtime_overrides_json",
                        overtimeOverridesJson.toAbsolutePath().normalize().toString());
            }
            ObjectNode summary = JSON.createObjectNode();
            OverrideSummary s =
                    overrideSummary != null ? overrideSummary : OverrideSummary.empty();
            summary.put("work_on", s.workOn());
            summary.put("work_off", s.workOff());
            summary.put("overtime_cells", s.overtimeCells());
            root.set("override_summary", summary);
            root.put("applied_at", LocalDateTime.now().format(APPLIED_AT_FMT));
            ObjectNode entries = JSON.createObjectNode();
            if (snapshot != null) {
                for (Map.Entry<String, Double> e : snapshot.entrySet()) {
                    entries.put(e.getKey(), e.getValue());
                }
            }
            root.set("entries", entries);
            Files.createDirectories(sidecar.getParent());
            JSON.writerWithDefaultPrettyPrinter().writeValue(sidecar.toFile(), root);
        } catch (Exception ignored) {
        }
    }

    public static Map<String, Double> tryLoadEntries(Path mainResultDispatchJson) {
        return tryLoadMeta(mainResultDispatchJson).entries();
    }

    public static Stage21TrialMeta tryLoadMeta(Path mainResultDispatchJson) {
        Path sidecar = sidecarPathFor(mainResultDispatchJson);
        if (sidecar == null || !Files.isRegularFile(sidecar)) {
            return Stage21TrialMeta.empty();
        }
        try {
            JsonNode root = JSON.readTree(Files.readString(sidecar, StandardCharsets.UTF_8));
            String expected =
                    mainResultDispatchJson != null
                            ? mainResultDispatchJson.toAbsolutePath().normalize().toString()
                            : "";
            String stored = root.path("source_json").asText("");
            if (!expected.isEmpty() && !stored.isEmpty() && !expected.equals(stored)) {
                return Stage21TrialMeta.empty();
            }
            boolean stage21Applied = root.path("stage21_applied").asBoolean(false);
            String stage21Json = root.path("stage21_result_dispatch_json").asText("");
            String overridesJson = root.path("overtime_overrides_json").asText("");
            OverrideSummary summary = parseOverrideSummary(root.path("override_summary"));
            String appliedAt = root.path("applied_at").asText("");
            Map<String, Double> entries = parseEntries(root.get("entries"));
            if (!stage21Applied && entries.isEmpty()) {
                return Stage21TrialMeta.empty();
            }
            return new Stage21TrialMeta(
                    stage21Applied || !entries.isEmpty(),
                    stage21Json,
                    overridesJson,
                    summary,
                    appliedAt,
                    entries);
        } catch (Exception ignored) {
            return Stage21TrialMeta.empty();
        }
    }

    private static OverrideSummary parseOverrideSummary(JsonNode node) {
        if (node == null || !node.isObject()) {
            return OverrideSummary.empty();
        }
        return new OverrideSummary(
                node.path("work_on").asInt(0),
                node.path("work_off").asInt(0),
                node.path("overtime_cells").asInt(0));
    }

    private static Map<String, Double> parseEntries(JsonNode entries) {
        if (entries == null || !entries.isObject()) {
            return Map.of();
        }
        Map<String, Double> out = new HashMap<>();
        entries.fields()
                .forEachRemaining(
                        f -> {
                            double v = f.getValue().asDouble(0.0);
                            if (v > 1e-3) {
                                out.put(f.getKey(), v);
                            }
                        });
        return Map.copyOf(out);
    }

    public static void deleteSidecar(Path mainResultDispatchJson) {
        Path sidecar = sidecarPathFor(mainResultDispatchJson);
        if (sidecar == null) {
            return;
        }
        try {
            Files.deleteIfExists(sidecar);
        } catch (Exception ignored) {
        }
    }
}
