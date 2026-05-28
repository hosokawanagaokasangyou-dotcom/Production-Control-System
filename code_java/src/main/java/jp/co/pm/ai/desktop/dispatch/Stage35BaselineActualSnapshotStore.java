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
 * 段階3.5 試行直前の {@link ResultDispatchSchema#COL_DISPATCH_QTY_ACTUAL} 合計（プロファイル×配台日）と
 * 勤怠上書きメタ。配台計画手動修正タブの (段階3後)/(段階3.5後) 行・勤怠適用バナー用。
 * 結果_配台表.json 隣の sidecar に永続化する。
 */
public final class Stage35BaselineActualSnapshotStore {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final DateTimeFormatter APPLIED_AT_FMT = DateTimeFormatter.ISO_LOCAL_DATE_TIME;

    private static final String SIDEcar_SUFFIX = ".stage35_baseline_actual.json";

    private Stage35BaselineActualSnapshotStore() {}

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

    public record Stage35TrialMeta(
            boolean stage35Applied,
            String overtimeOverridesJson,
            OverrideSummary overrideSummary,
            String appliedAt,
            Map<String, Double> entries) {

        public static Stage35TrialMeta empty() {
            return new Stage35TrialMeta(
                    false, "", OverrideSummary.empty(), "", Map.of());
        }

        public boolean hasTrialApplied() {
            return stage35Applied || (entries != null && !entries.isEmpty());
        }

        public Path overtimeOverridesPath() {
            if (overtimeOverridesJson == null || overtimeOverridesJson.isBlank()) {
                return null;
            }
            return Path.of(overtimeOverridesJson);
        }
    }

    public static Path sidecarPathFor(Path resultDispatchJson) {
        if (resultDispatchJson == null) {
            return null;
        }
        String fileName = resultDispatchJson.getFileName().toString();
        if (!fileName.endsWith(".json")) {
            return resultDispatchJson.resolveSibling(fileName + SIDEcar_SUFFIX);
        }
        return resultDispatchJson.resolveSibling(
                fileName.substring(0, fileName.length() - 5) + SIDEcar_SUFFIX);
    }

    public static Map<String, Double> captureFromDocument(
            ResultDispatchDocument sourceDoc, List<LocalDate> axis) {
        Map<String, Double> out = new LinkedHashMap<>();
        if (sourceDoc == null || axis == null || axis.isEmpty()) {
            return out;
        }
        if (!sourceDoc.columns().contains(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL)) {
            return out;
        }
        List<Map<String, String>> profiles =
                ResultDispatchPivot.distinctWideTaskProfiles(
                        sourceDoc.columns(),
                        sourceDoc.rows(),
                        ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
        for (Map<String, String> profile : profiles) {
            for (LocalDate day : axis) {
                double actual =
                        ResultDispatchPivot.sumActualQuantityForProfileAndDateForWideMerge(
                                sourceDoc.rows(),
                                profile,
                                day,
                                ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
                String key = cellKey(profile, day);
                if (actual > 1e-3) {
                    out.put(key, actual);
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

    /** @deprecated {@link #writeWithMeta(Path, Map, Path, OverrideSummary)} を使用 */
    @Deprecated
    public static void write(Path resultDispatchJson, Map<String, Double> snapshot) {
        writeWithMeta(
                resultDispatchJson,
                snapshot,
                null,
                OverrideSummary.empty());
    }

    public static void writeWithMeta(
            Path resultDispatchJson,
            Map<String, Double> snapshot,
            Path overtimeOverridesJson,
            OverrideSummary overrideSummary) {
        Path sidecar = sidecarPathFor(resultDispatchJson);
        if (sidecar == null) {
            return;
        }
        try {
            ObjectNode root = JSON.createObjectNode();
            root.put("source_json", resultDispatchJson.toAbsolutePath().normalize().toString());
            root.put("stage35_applied", true);
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

    /** 旧 sidecar（entries のみ）互換。 */
    public static Map<String, Double> tryLoad(Path resultDispatchJson) {
        return tryLoadMeta(resultDispatchJson).entries();
    }

    public static Stage35TrialMeta tryLoadMeta(Path resultDispatchJson) {
        Path sidecar = sidecarPathFor(resultDispatchJson);
        if (sidecar == null || !Files.isRegularFile(sidecar)) {
            return Stage35TrialMeta.empty();
        }
        try {
            JsonNode root = JSON.readTree(Files.readString(sidecar, StandardCharsets.UTF_8));
            String expected =
                    resultDispatchJson != null
                            ? resultDispatchJson.toAbsolutePath().normalize().toString()
                            : "";
            String stored = root.path("source_json").asText("");
            if (!expected.isEmpty() && !stored.isEmpty() && !expected.equals(stored)) {
                return Stage35TrialMeta.empty();
            }
            boolean stage35Applied = root.path("stage35_applied").asBoolean(false);
            String overridesJson = root.path("overtime_overrides_json").asText("");
            OverrideSummary summary = parseOverrideSummary(root.path("override_summary"));
            String appliedAt = root.path("applied_at").asText("");
            Map<String, Double> entries = parseEntries(root.get("entries"));
            if (!stage35Applied && entries.isEmpty()) {
                return Stage35TrialMeta.empty();
            }
            return new Stage35TrialMeta(
                    stage35Applied || !entries.isEmpty(),
                    overridesJson,
                    summary,
                    appliedAt,
                    entries);
        } catch (Exception ignored) {
            return Stage35TrialMeta.empty();
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

    public static void deleteSidecar(Path resultDispatchJson) {
        Path sidecar = sidecarPathFor(resultDispatchJson);
        if (sidecar == null) {
            return;
        }
        try {
            Files.deleteIfExists(sidecar);
        } catch (Exception ignored) {
        }
    }
}
