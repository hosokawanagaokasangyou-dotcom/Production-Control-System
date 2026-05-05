package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * Persists last-used paths under {@code ~/.pm-ai-desktop/session-state.json} so tabs reload the same files on
 * the next launch.
 */
public final class DesktopSessionStateStore {

    private static final ObjectMapper JSON = new ObjectMapper();
    private static final Path STORE =
            Paths.get(System.getProperty("user.home"), ".pm-ai-desktop", "session-state.json");

    private DesktopSessionStateStore() {}

    public static DesktopSessionState load() {
        try {
            if (!Files.isRegularFile(STORE)) {
                return DesktopSessionState.empty();
            }
            JsonNode root = JSON.readTree(STORE.toFile());
            if (root == null || !root.isObject()) {
                return DesktopSessionState.empty();
            }
            return new DesktopSessionState(
                    text(root, "planInputPath"),
                    text(root, "planInputSheet"),
                    text(root, "stage1PreviewPath"),
                    text(root, "stage1PreviewSheet"),
                    text(root, "excludeRulesPath"),
                    text(root, "mainRunWorkbook"),
                    text(root, "mainRunPythonExe"),
                    text(root, "mainRunScriptDir"),
                    optionalDouble(root, "windowWidth", 0d),
                    optionalDouble(root, "windowHeight", 0d),
                    optionalDouble(root, "windowX", Double.NaN),
                    optionalDouble(root, "windowY", Double.NaN),
                    text(root, "uiTheme"),
                    text(root, "logFontFamily"),
                    optionalDouble(root, "logFontSize", 0d),
                    text(root, "mainRunLogFilter"),
                    loadStringList(root, "mainRunLogLines"),
                    optionalDouble(root, "mainRunLogScroll", Double.NaN),
                    text(root, "mainRunStage2ProductionPlan"),
                    text(root, "mainRunStage2MemberSchedule"),
                    optionalBoolean(root, "mainRunStage2WriteExcel", true),
                    text(root, "mainRunStage2ResultBookFont"),
                    loadUiEnvRows(root),
                    loadStringList(root, "mainShellTabOrder"),
                    optionalDouble(root, "equipmentGanttGraphicZoomPercent", 0d),
                    optionalDouble(root, "equipmentGanttDateColWidth", 0d),
                    optionalDouble(root, "equipmentGanttMachineColWidth", 0d),
                    optionalDouble(root, "equipmentGanttProcessColWidth", 0d),
                    text(root, "equipmentGanttBarFontFamily"),
                    optionalDouble(root, "equipmentGanttBarFontPercent", 0d),
                    optionalDouble(root, "equipmentGanttRowHeightPercent", 0d),
                    optionalDouble(root, "equipmentGanttHeaderHeightPercent", 0d),
                    optionalDouble(root, "equipmentGanttSlotWidthPercent", 0d),
                    optionalDouble(root, "equipmentGanttShiftWheelHScrollPercent", 0d),
                    optionalBoolean(root, "equipmentGanttPersonBadgeEnabled", true),
                    text(root, "equipmentGanttPersonBadgeFontFamily"),
                    optionalDouble(root, "equipmentGanttPersonBadgeFontPercent", 0d),
                    text(root, "equipmentGanttPersonBadgeFillHex"),
                    text(root, "equipmentGanttPersonBadgeTextHex"),
                    text(root, "equipmentGanttPersonBadgeStrokeHex"),
                    optionalDouble(root, "equipmentGanttPersonBadgeStrokeWidth", -1d),
                    optionalDouble(root, "equipmentGanttPersonBadgeCornerRadius", -1d),
                    optionalBoolean(root, "equipmentGanttPersonBadgePill", false),
                    text(root, "equipmentGanttPersonBadgeGlowColorHex"),
                    optionalDouble(root, "equipmentGanttPersonBadgeGlowRadius", -1d),
                    optionalDouble(root, "equipmentGanttPersonBadgeGlowSpread", -1d),
                    loadPersonBadgeStyleMap(root, "equipmentGanttPersonBadgeStylesByLabel"),
                    loadPersonBadgeStyleMap(root, "equipmentGanttPersonBadgeStylesByMemberKey"));
        } catch (IOException e) {
            return DesktopSessionState.empty();
        }
    }

    public static void save(DesktopSessionState state) {
        try {
            Files.createDirectories(STORE.getParent());
            ObjectNode root = JSON.createObjectNode();
            put(root, "planInputPath", state.planInputPath());
            put(root, "planInputSheet", state.planInputSheet());
            put(root, "stage1PreviewPath", state.stage1PreviewPath());
            put(root, "stage1PreviewSheet", state.stage1PreviewSheet());
            put(root, "excludeRulesPath", state.excludeRulesPath());
            put(root, "mainRunWorkbook", state.mainRunWorkbook());
            put(root, "mainRunPythonExe", state.mainRunPythonExe());
            put(root, "mainRunScriptDir", state.mainRunScriptDir());
            put(root, "uiTheme", state.uiTheme());
            put(root, "logFontFamily", state.logFontFamily());
            putLogFontSize(root, state.logFontSize());
            put(root, "mainRunLogFilter", state.mainRunLogFilter());
            putMainRunLogLines(root, state.mainRunLogLines());
            putMainRunLogScroll(root, state.mainRunLogScroll());
            put(root, "mainRunStage2ProductionPlan", state.mainRunStage2ProductionPlan());
            put(root, "mainRunStage2MemberSchedule", state.mainRunStage2MemberSchedule());
            root.put("mainRunStage2WriteExcel", state.mainRunStage2WriteExcel());
            put(root, "mainRunStage2ResultBookFont", state.mainRunStage2ResultBookFont());
            putUiEnvRows(root, state.uiEnvRows());
            putMainShellTabOrder(root, state.mainShellTabOrder());
            putEquipmentGanttGraphicPrefs(root, state);
            putWindowGeometry(root, state);
            JSON.writerWithDefaultPrettyPrinter().writeValue(STORE.toFile(), root);
        } catch (IOException ignored) {
        }
    }

    private static String text(JsonNode root, String key) {
        JsonNode n = root.get(key);
        if (n == null || n.isNull() || !n.isTextual()) {
            return "";
        }
        return n.asText("");
    }

    private static double optionalDouble(JsonNode root, String key, double defaultValue) {
        JsonNode n = root.get(key);
        if (n == null || n.isNull() || !n.isNumber()) {
            return defaultValue;
        }
        return n.asDouble();
    }

    private static boolean optionalBoolean(JsonNode root, String key, boolean defaultValue) {
        JsonNode n = root.get(key);
        if (n == null || n.isNull()) {
            return defaultValue;
        }
        if (n.isBoolean()) {
            return n.booleanValue();
        }
        if (n.isTextual()) {
            String t = n.asText("").trim().toLowerCase(Locale.ROOT);
            return !List.of("0", "false", "no", "off", "none").contains(t);
        }
        return defaultValue;
    }

    private static void put(ObjectNode root, String key, String value) {
        if (value != null && !value.isBlank()) {
            root.put(key, value.trim());
        }
    }

    private static void putLogFontSize(ObjectNode root, double sizePoints) {
        if (Double.isFinite(sizePoints) && sizePoints > 0) {
            root.put("logFontSize", sizePoints);
        }
    }

    private static List<String> loadStringList(JsonNode root, String key) {
        JsonNode arr = root.get(key);
        if (arr == null || !arr.isArray()) {
            return List.of();
        }
        List<String> out = new ArrayList<>();
        for (JsonNode el : arr) {
            if (el != null && el.isTextual()) {
                out.add(el.asText(""));
            } else if (el != null && el.isValueNode()) {
                out.add(el.asText(""));
            }
        }
        return List.copyOf(out);
    }

    private static void putMainRunLogLines(ObjectNode root, List<String> lines) {
        if (lines == null || lines.isEmpty()) {
            return;
        }
        ArrayNode arr = JSON.createArrayNode();
        for (String s : lines) {
            arr.add(s != null ? s : "");
        }
        root.set("mainRunLogLines", arr);
    }

    private static void putMainRunLogScroll(ObjectNode root, double scroll) {
        if (Double.isFinite(scroll)) {
            root.put("mainRunLogScroll", scroll);
        }
    }

    private static List<UiEnvRowSnapshot> loadUiEnvRows(JsonNode root) {
        JsonNode arr = root.get("uiEnvRows");
        if (arr == null || !arr.isArray()) {
            return List.of();
        }
        List<UiEnvRowSnapshot> out = new ArrayList<>();
        for (JsonNode el : arr) {
            if (el != null && el.isObject()) {
                out.add(
                        new UiEnvRowSnapshot(
                                text(el, "name"), text(el, "value"), text(el, "description")));
            }
        }
        return List.copyOf(out);
    }

    private static void putUiEnvRows(ObjectNode root, List<UiEnvRowSnapshot> rows) {
        ArrayNode arr = JSON.createArrayNode();
        if (rows != null) {
            for (UiEnvRowSnapshot r : rows) {
                ObjectNode o = arr.addObject();
                o.put("name", r.name() != null ? r.name() : "");
                o.put("value", r.value() != null ? r.value() : "");
                o.put("description", r.description() != null ? r.description() : "");
            }
        }
        root.set("uiEnvRows", arr);
    }

    private static void putMainShellTabOrder(ObjectNode root, List<String> order) {
        if (order == null || order.isEmpty()) {
            return;
        }
        ArrayNode arr = JSON.createArrayNode();
        for (String s : order) {
            if (s != null && !s.isBlank()) {
                arr.add(s.trim());
            }
        }
        if (!arr.isEmpty()) {
            root.set("mainShellTabOrder", arr);
        }
    }

    private static void putEquipmentGanttGraphicPrefs(ObjectNode root, DesktopSessionState state) {
        double z = state.equipmentGanttGraphicZoomPercent();
        if (Double.isFinite(z) && z >= 50 && z <= 200) {
            root.put("equipmentGanttGraphicZoomPercent", z);
        }
        double dw = state.equipmentGanttDateColWidth();
        if (Double.isFinite(dw) && dw > 0) {
            root.put("equipmentGanttDateColWidth", dw);
        }
        double mw = state.equipmentGanttMachineColWidth();
        if (Double.isFinite(mw) && mw > 0) {
            root.put("equipmentGanttMachineColWidth", mw);
        }
        double pw = state.equipmentGanttProcessColWidth();
        if (Double.isFinite(pw) && pw > 0) {
            root.put("equipmentGanttProcessColWidth", pw);
        }
        String bf = state.equipmentGanttBarFontFamily();
        if (bf != null && !bf.isBlank()) {
            root.put("equipmentGanttBarFontFamily", bf.strip());
        }
        double bfp = state.equipmentGanttBarFontPercent();
        if (Double.isFinite(bfp) && bfp >= 50 && bfp <= 200) {
            root.put("equipmentGanttBarFontPercent", bfp);
        }
        double rh = state.equipmentGanttRowHeightPercent();
        if (Double.isFinite(rh) && rh >= 50 && rh <= 200) {
            root.put("equipmentGanttRowHeightPercent", rh);
        }
        double hh = state.equipmentGanttHeaderHeightPercent();
        if (Double.isFinite(hh) && hh >= 50 && hh <= 200) {
            root.put("equipmentGanttHeaderHeightPercent", hh);
        }
        double sw = state.equipmentGanttSlotWidthPercent();
        if (Double.isFinite(sw) && sw >= 50 && sw <= 500) {
            root.put("equipmentGanttSlotWidthPercent", sw);
        }
        double sh = state.equipmentGanttShiftWheelHScrollPercent();
        if (Double.isFinite(sh) && sh >= 50 && sh <= 1000) {
            root.put("equipmentGanttShiftWheelHScrollPercent", sh);
        }
        root.put("equipmentGanttPersonBadgeEnabled", state.equipmentGanttPersonBadgeEnabled());
        put(root, "equipmentGanttPersonBadgeFontFamily", state.equipmentGanttPersonBadgeFontFamily());
        double bpf = state.equipmentGanttPersonBadgeFontPercent();
        if (Double.isFinite(bpf) && bpf > 0 && bpf <= 300) {
            root.put("equipmentGanttPersonBadgeFontPercent", bpf);
        }
        put(root, "equipmentGanttPersonBadgeFillHex", state.equipmentGanttPersonBadgeFillHex());
        put(root, "equipmentGanttPersonBadgeTextHex", state.equipmentGanttPersonBadgeTextHex());
        put(root, "equipmentGanttPersonBadgeStrokeHex", state.equipmentGanttPersonBadgeStrokeHex());
        double stw = state.equipmentGanttPersonBadgeStrokeWidth();
        if (Double.isFinite(stw) && stw >= 0) {
            root.put("equipmentGanttPersonBadgeStrokeWidth", stw);
        }
        double cr = state.equipmentGanttPersonBadgeCornerRadius();
        if (Double.isFinite(cr) && cr >= 0) {
            root.put("equipmentGanttPersonBadgeCornerRadius", cr);
        }
        root.put("equipmentGanttPersonBadgePill", state.equipmentGanttPersonBadgePill());
        put(root, "equipmentGanttPersonBadgeGlowColorHex", state.equipmentGanttPersonBadgeGlowColorHex());
        double gr = state.equipmentGanttPersonBadgeGlowRadius();
        if (Double.isFinite(gr) && gr >= 0) {
            root.put("equipmentGanttPersonBadgeGlowRadius", gr);
        }
        double gs = state.equipmentGanttPersonBadgeGlowSpread();
        if (Double.isFinite(gs) && gs >= 0 && gs <= 1) {
            root.put("equipmentGanttPersonBadgeGlowSpread", gs);
        }
        putPersonBadgeStyleMap(root, state.equipmentGanttPersonBadgeStylesByLabel(), "equipmentGanttPersonBadgeStylesByLabel");
        putPersonBadgeStyleMap(root, state.equipmentGanttPersonBadgeStylesByMemberKey(), "equipmentGanttPersonBadgeStylesByMemberKey");
    }

    private static Map<String, PersonBadgeStyle> loadPersonBadgeStyleMap(JsonNode root, String jsonKey) {
        JsonNode obj = root.get(jsonKey);
        if (obj == null || !obj.isObject()) {
            return Map.of();
        }
        Map<String, PersonBadgeStyle> out = new LinkedHashMap<>();
        for (Iterator<String> it = obj.fieldNames(); it.hasNext(); ) {
            String field = it.next();
            JsonNode el = obj.get(field);
            if (el != null && el.isObject()) {
                PersonBadgeStyle st = loadPersonBadgeStyleObject(el);
                String k = PersonBadgeStyle.normalizeLabelKey(field);
                if (st != null && !k.isEmpty()) {
                    out.put(k, st);
                }
            }
        }
        return Map.copyOf(out);
    }

    private static PersonBadgeStyle loadPersonBadgeStyleObject(JsonNode o) {
        PersonBadgeStyle d = PersonBadgeStyle.defaultStyle();
        String ff = text(o, "fontFamily");
        double fp = optionalDouble(o, "fontPercent", 0d);
        String fill = text(o, "fillHex");
        String tx = text(o, "textHex");
        String stk = text(o, "strokeHex");
        double stw = optionalDouble(o, "strokeWidth", -1d);
        double cr = optionalDouble(o, "cornerRadius", -1d);
        boolean pill = optionalBoolean(o, "pill", false);
        String glow = text(o, "glowColorHex");
        double gr = optionalDouble(o, "glowRadius", -1d);
        double gs = optionalDouble(o, "glowSpread", -1d);
        return new PersonBadgeStyle(
                ff,
                fp > 0 && fp <= 300 ? fp : d.fontPercent(),
                nzStored(fill, d.fillHex()),
                nzStored(tx, d.textHex()),
                nzStored(stk, d.strokeHex()),
                stw >= 0 ? stw : d.strokeWidth(),
                cr >= 0 ? cr : d.cornerRadius(),
                pill,
                nzStored(glow, d.glowColorHex()),
                gr >= 0 ? gr : d.glowRadius(),
                gs >= 0 && gs <= 1 ? gs : d.glowSpread());
    }

    private static String nzStored(String s, String def) {
        return s != null && !s.isBlank() ? s.strip() : def;
    }

    private static void putPersonBadgeStyleMap(
            ObjectNode root, Map<String, PersonBadgeStyle> styleMap, String jsonKey) {
        Map<String, PersonBadgeStyle> m = styleMap;
        if (m == null || m.isEmpty()) {
            return;
        }
        ObjectNode bag = JSON.createObjectNode();
        for (Map.Entry<String, PersonBadgeStyle> e : m.entrySet()) {
            if (e.getKey() == null || e.getKey().isBlank() || e.getValue() == null) {
                continue;
            }
            String canon = PersonBadgeStyle.normalizeLabelKey(e.getKey());
            if (canon.isEmpty()) {
                continue;
            }
            PersonBadgeStyle st = e.getValue();
            ObjectNode o = bag.putObject(canon);
            o.put("fontFamily", st.fontFamily() != null ? st.fontFamily() : "");
            o.put("fontPercent", st.fontPercent());
            o.put("fillHex", st.fillHex());
            o.put("textHex", st.textHex());
            o.put("strokeHex", st.strokeHex());
            o.put("strokeWidth", st.strokeWidth());
            o.put("cornerRadius", st.cornerRadius());
            o.put("pill", st.pill());
            o.put("glowColorHex", st.glowColorHex());
            o.put("glowRadius", st.glowRadius());
            o.put("glowSpread", st.glowSpread());
        }
        if (!bag.isEmpty()) {
            root.set(jsonKey, bag);
        }
    }

    private static void putWindowGeometry(ObjectNode root, DesktopSessionState state) {
        double w = state.windowWidth();
        double h = state.windowHeight();
        if (Double.isFinite(w) && w > 0) {
            root.put("windowWidth", w);
        }
        if (Double.isFinite(h) && h > 0) {
            root.put("windowHeight", h);
        }
        double x = state.windowX();
        double y = state.windowY();
        if (Double.isFinite(x)) {
            root.put("windowX", x);
        }
        if (Double.isFinite(y)) {
            root.put("windowY", y);
        }
    }
}
