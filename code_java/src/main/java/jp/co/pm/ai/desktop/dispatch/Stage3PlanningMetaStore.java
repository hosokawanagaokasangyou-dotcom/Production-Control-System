package jp.co.pm.ai.desktop.dispatch;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.LinkedHashMap;
import java.util.Map;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

/** 段階3.0/3.1/3.2 実行後の計画種別（{@code 結果_配台表.json} 横の sidecar）。 */
public final class Stage3PlanningMetaStore {

    private static final ObjectMapper JSON = new ObjectMapper();
    private static final String SIDECAR_SUFFIX = ".stage3_planning_meta.json";
    private static final String FIELD_VARIANT = "variant";
    private static final String FIELD_APPLIED_AT = "appliedAtEpochMs";
    private static final String FIELD_BASELINE_ENTRIES = "baselineEntries";
    private static final DateTimeFormatter DELIVERY_CAL_DATE_FMT =
            DateTimeFormatter.ofPattern("yyyy/MM/dd");

    public enum Variant {
        STAGE3_0("3.0"),
        STAGE3_1("3.1"),
        STAGE3_2("3.2");

        private final String id;

        Variant(String id) {
            this.id = id;
        }

        public String id() {
            return id;
        }

        /** 日付セル内の実績行ラベル（例: {@code (段階3.0後)}）。 */
        public String actualQtyLabel() {
            return "(段階" + id + "後)";
        }

        /** 日付セル内の手修正行ラベル（例: {@code (段階3.0改)}）。 */
        public String revisedQtyLabel() {
            return "(段階" + id + "改)";
        }

        /** 配台表バッジ（例: {@code 段階3.0}）。 */
        public String badgeText() {
            return "段階" + id;
        }

        public static Variant fromId(String raw) {
            if (raw == null || raw.isBlank()) {
                return null;
            }
            String t = raw.trim();
            for (Variant v : values()) {
                if (v.id.equals(t)) {
                    return v;
                }
            }
            return null;
        }
    }

    private Stage3PlanningMetaStore() {}

    public static Path sidecarPath(Path dispatchJson) {
        if (dispatchJson == null) {
            return null;
        }
        String name = dispatchJson.getFileName().toString();
        return dispatchJson.resolveSibling(name + SIDECAR_SUFFIX);
    }

    public static void writeVariant(Path dispatchJson, Variant variant) {
        if (dispatchJson == null || variant == null) {
            return;
        }
        writeSidecar(dispatchJson, variant, readBaselineEntries(dispatchJson));
    }

    /** 段階3.0/3.1/3.2 実行直前: (段階3前) 比較用 baseline を sidecar へ保存する。 */
    public static void writeBaselineEntries(Path dispatchJson, Map<String, Double> entries) {
        if (dispatchJson == null || entries == null || entries.isEmpty()) {
            return;
        }
        Variant existingVariant = readVariant(dispatchJson);
        writeSidecar(dispatchJson, existingVariant, entries);
    }

    public static Map<String, Double> readBaselineEntries(Path dispatchJson) {
        if (dispatchJson == null) {
            return Map.of();
        }
        Path sidecar = sidecarPath(dispatchJson);
        if (sidecar == null || !Files.isRegularFile(sidecar)) {
            return Map.of();
        }
        try {
            JsonNode root = JSON.readTree(sidecar.toFile());
            JsonNode baseline = root.get(FIELD_BASELINE_ENTRIES);
            if (baseline == null || !baseline.isObject()) {
                return Map.of();
            }
            Map<String, Double> out = new LinkedHashMap<>();
            baseline.fields()
                    .forEachRemaining(
                            e -> {
                                JsonNode v = e.getValue();
                                if (v != null && v.isNumber()) {
                                    out.put(e.getKey(), v.asDouble());
                                }
                            });
            return Map.copyOf(out);
        } catch (Exception ignored) {
            return Map.of();
        }
    }

    /**
     * 納期管理ビュー用: baseline sidecar を mk→tid→yyyy/MM/dd ルックアップへ変換する。
     * キー形式は {@link DispatchTrialShortages#wideShortfallKey}。
     */
    public static Map<String, Map<String, Map<String, Double>>> buildBaselineDispatchLookup(
            Map<String, Double> baselineEntries) {
        if (baselineEntries == null || baselineEntries.isEmpty()) {
            return Map.of();
        }
        Map<String, Map<String, Map<String, Double>>> result = new LinkedHashMap<>();
        for (Map.Entry<String, Double> e : baselineEntries.entrySet()) {
            String[] parts = e.getKey().split("\u0001", -1);
            if (parts.length < 3) {
                continue;
            }
            String tid = parts[0];
            String mk = parts[1];
            String isoDate = parts[2];
            if (tid.isEmpty() || mk.isEmpty() || isoDate.isEmpty()) {
                continue;
            }
            String deliveryDate;
            try {
                deliveryDate = LocalDate.parse(isoDate).format(DELIVERY_CAL_DATE_FMT);
            } catch (Exception ex) {
                continue;
            }
            double qty = e.getValue() != null ? e.getValue() : 0.0;
            if (Math.abs(qty) <= 1e-12) {
                continue;
            }
            result.computeIfAbsent(mk, k -> new LinkedHashMap<>())
                    .computeIfAbsent(tid, k -> new LinkedHashMap<>())
                    .merge(deliveryDate, qty, Double::sum);
        }
        return result;
    }

    public static boolean hasPipelinePlanningVariant(Path dispatchJson) {
        return readVariant(dispatchJson) != null;
    }

    private static Variant readVariant(Path dispatchJson) {
        if (dispatchJson == null || !Files.isRegularFile(dispatchJson)) {
            return null;
        }
        Path sidecar = sidecarPath(dispatchJson);
        if (sidecar == null || !Files.isRegularFile(sidecar)) {
            return null;
        }
        try {
            JsonNode root = JSON.readTree(sidecar.toFile());
            return Variant.fromId(text(root, FIELD_VARIANT));
        } catch (Exception ignored) {
            return null;
        }
    }

    private static void writeSidecar(
            Path dispatchJson, Variant variant, Map<String, Double> baselineEntries) {
        Path sidecar = sidecarPath(dispatchJson);
        if (sidecar == null) {
            return;
        }
        try {
            ObjectNode root = JSON.createObjectNode();
            if (variant != null) {
                root.put(FIELD_VARIANT, variant.id());
                root.put(FIELD_APPLIED_AT, Instant.now().toEpochMilli());
            }
            if (baselineEntries != null && !baselineEntries.isEmpty()) {
                ObjectNode baseline = JSON.createObjectNode();
                for (Map.Entry<String, Double> e : baselineEntries.entrySet()) {
                    baseline.put(e.getKey(), e.getValue());
                }
                root.set(FIELD_BASELINE_ENTRIES, baseline);
            }
            if (sidecar.getParent() != null) {
                Files.createDirectories(sidecar.getParent());
            }
            JSON.writerWithDefaultPrettyPrinter().writeValue(sidecar.toFile(), root);
        } catch (Exception ignored) {
            // sidecar 失敗で配台本体は止めない
        }
    }

    /** sidecar が無いが実配台数量列がある旧 JSON は {@link ResultDispatchStage3Support.Stage3PlanningVariant#LEGACY}。 */
    public static ResultDispatchStage3Support.Stage3PlanningVariant readPlanningVariant(Path dispatchJson) {
        if (dispatchJson == null || !Files.isRegularFile(dispatchJson)) {
            return ResultDispatchStage3Support.Stage3PlanningVariant.NONE;
        }
        Path sidecar = sidecarPath(dispatchJson);
        if (sidecar != null && Files.isRegularFile(sidecar)) {
            try {
                JsonNode root = JSON.readTree(sidecar.toFile());
                Variant v = Variant.fromId(text(root, FIELD_VARIANT));
                if (v != null) {
                    return ResultDispatchStage3Support.Stage3PlanningVariant.fromMetaVariant(v);
                }
            } catch (Exception ignored) {
                // fall through
            }
        }
        if (ResultDispatchStage3Support.detectStage3FromDispatchJsonPath(dispatchJson)) {
            return ResultDispatchStage3Support.Stage3PlanningVariant.LEGACY;
        }
        return ResultDispatchStage3Support.Stage3PlanningVariant.NONE;
    }

    /** 段階2再実行などで古い段階3メタが残らないよう sidecar を削除する。 */
    public static void deleteSidecar(Path dispatchJson) {
        Path sidecar = sidecarPath(dispatchJson);
        if (sidecar == null) {
            return;
        }
        try {
            Files.deleteIfExists(sidecar);
        } catch (Exception ignored) {
        }
    }

    private static String text(JsonNode node, String field) {
        JsonNode v = node != null ? node.get(field) : null;
        return v != null && !v.isNull() ? v.asText("") : "";
    }
}
