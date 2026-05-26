package jp.co.pm.ai.desktop.dispatch;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * 段階3.5 試行直前の {@link ResultDispatchSchema#COL_DISPATCH_QTY_ACTUAL} 合計（プロファイル×配台日）。
 * 配台計画手動修正タブの (段階3後) 行表示用。結果_配台表.json 隣の sidecar に永続化する。
 */
public final class Stage35BaselineActualSnapshotStore {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final String SIDEcar_SUFFIX = ".stage35_baseline_actual.json";

    private Stage35BaselineActualSnapshotStore() {}

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

    public static void write(Path resultDispatchJson, Map<String, Double> snapshot) {
        Path sidecar = sidecarPathFor(resultDispatchJson);
        if (sidecar == null) {
            return;
        }
        try {
            ObjectNode root = JSON.createObjectNode();
            root.put("source_json", resultDispatchJson.toAbsolutePath().normalize().toString());
            ObjectNode entries = JSON.createObjectNode();
            if (snapshot != null) {
                for (Map.Entry<String, Double> e : snapshot.entrySet()) {
                    entries.put(e.getKey(), e.getValue());
                }
            }
            root.set("entries", entries);
            Files.createDirectories(sidecar.getParent());
            JSON.writerWithDefaultPrettyPrinter()
                    .writeValue(sidecar.toFile(), root);
        } catch (Exception ignored) {
        }
    }

    public static Map<String, Double> tryLoad(Path resultDispatchJson) {
        Path sidecar = sidecarPathFor(resultDispatchJson);
        if (sidecar == null || !Files.isRegularFile(sidecar)) {
            return Map.of();
        }
        try {
            JsonNode root = JSON.readTree(Files.readString(sidecar, StandardCharsets.UTF_8));
            String expected =
                    resultDispatchJson != null
                            ? resultDispatchJson.toAbsolutePath().normalize().toString()
                            : "";
            String stored = root.path("source_json").asText("");
            if (!expected.isEmpty() && !stored.isEmpty() && !expected.equals(stored)) {
                return Map.of();
            }
            JsonNode entries = root.get("entries");
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
        } catch (Exception ignored) {
            return Map.of();
        }
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
