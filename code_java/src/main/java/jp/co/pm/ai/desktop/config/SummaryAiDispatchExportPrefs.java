package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import jp.co.pm.ai.desktop.io.SummaryAiDispatchWorkbookExporter;

/**
 * {@link SummaryAiDispatchWorkbookExporter} 向けの Excel 出力カスタマイズ（見出し列数・フォント・非日付列順・テーマ）。
 */
public final class SummaryAiDispatchExportPrefs {

    public static final String DEFAULT_FONT_FAMILY = "BIZ UDP\u30b4\u30b7\u30c3\u30af";
    public static final int DEFAULT_FONT_SIZE_PT = 11;
    public static final DesktopTheme DEFAULT_THEME = DesktopTheme.LIGHT;

    private static final ObjectMapper JSON = new ObjectMapper();
    private static final Path DEFAULT_STORE =
            Paths.get(System.getProperty("user.home"), ".pm-ai-desktop", "summary-ai-dispatch-export.json");

    /** テスト用: {@code -Dpm.ai.test.summaryExportPrefs=/path/to.json} */
    private static final String TEST_STORE_PROPERTY = "pm.ai.test.summaryExportPrefs";

    public enum SheetKey {
        MAIN_COMPARE(
                "mainCompare",
                SummaryAiDispatchWorkbookExporter.SHEET_MAIN_COMPARE,
                7),
        DISPATCH(
                "dispatch",
                SummaryAiDispatchWorkbookExporter.SHEET_DISPATCH,
                5),
        ACTUALS(
                "actuals",
                SummaryAiDispatchWorkbookExporter.SHEET_ACTUALS,
                0),
        ALADDIN(
                "aladdin",
                SummaryAiDispatchWorkbookExporter.SHEET_ALADDIN,
                7);

        private final String jsonKey;
        private final String sheetName;
        private final int defaultFrozenColumns;

        SheetKey(String jsonKey, String sheetName, int defaultFrozenColumns) {
            this.jsonKey = jsonKey;
            this.sheetName = sheetName;
            this.defaultFrozenColumns = defaultFrozenColumns;
        }

        public String jsonKey() {
            return jsonKey;
        }

        public String sheetName() {
            return sheetName;
        }

        public int defaultFrozenColumns() {
            return defaultFrozenColumns;
        }

        public static SheetKey forSheetName(String name) {
            if (name == null) {
                return null;
            }
            for (SheetKey k : values()) {
                if (k.sheetName.equals(name)) {
                    return k;
                }
            }
            return null;
        }
    }

    public record SheetPrefs(int frozenColumnCount, List<String> nonDateColumnOrder) {
        public SheetPrefs {
            frozenColumnCount = Math.max(0, Math.min(frozenColumnCount, 10_000));
            nonDateColumnOrder =
                    nonDateColumnOrder != null ? List.copyOf(nonDateColumnOrder) : List.of();
        }

        public static SheetPrefs defaults(SheetKey key) {
            return new SheetPrefs(key.defaultFrozenColumns(), List.of());
        }
    }

    public record ExportPrefs(
            String fontFamily,
            int fontSizePt,
            DesktopTheme theme,
            Map<SheetKey, SheetPrefs> sheets) {

        public ExportPrefs {
            fontFamily = fontFamily != null ? fontFamily : DEFAULT_FONT_FAMILY;
            fontSizePt = Math.max(6, Math.min(fontSizePt, 72));
            theme = theme != null ? theme : DEFAULT_THEME;
            Map<SheetKey, SheetPrefs> m = new LinkedHashMap<>();
            for (SheetKey k : SheetKey.values()) {
                SheetPrefs p = sheets != null ? sheets.get(k) : null;
                m.put(k, p != null ? p : SheetPrefs.defaults(k));
            }
            sheets = Map.copyOf(m);
        }

        public static ExportPrefs defaults() {
            return new ExportPrefs(DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE_PT, DEFAULT_THEME, Map.of());
        }

        public SheetPrefs sheet(SheetKey key) {
            return sheets.getOrDefault(key, SheetPrefs.defaults(key));
        }
    }

    private SummaryAiDispatchExportPrefs() {}

    public static Path storePath() {
        return resolveStore();
    }

    private static Path resolveStore() {
        String test = System.getProperty(TEST_STORE_PROPERTY);
        if (test != null && !test.isBlank()) {
            return Path.of(test.trim());
        }
        return DEFAULT_STORE;
    }

    public static ExportPrefs load() {
        Path store = resolveStore();
        try {
            if (!Files.isRegularFile(store)) {
                return ExportPrefs.defaults();
            }
            JsonNode root = JSON.readTree(store.toFile());
            if (root == null || !root.isObject()) {
                return ExportPrefs.defaults();
            }
            String font = root.path("fontFamily").asText(DEFAULT_FONT_FAMILY);
            int size = root.path("fontSizePt").asInt(DEFAULT_FONT_SIZE_PT);
            DesktopTheme theme =
                    DesktopTheme.fromStored(root.path("themeId").asText(DEFAULT_THEME.storedId()));
            if (theme == null) {
                theme = DEFAULT_THEME;
            }
            Map<SheetKey, SheetPrefs> sheets = new LinkedHashMap<>();
            JsonNode sheetsNode = root.get("sheets");
            for (SheetKey k : SheetKey.values()) {
                JsonNode sn =
                        sheetsNode != null && sheetsNode.isObject()
                                ? sheetsNode.get(k.jsonKey())
                                : null;
                int frozen =
                        sn != null && sn.isObject()
                                ? sn.path("frozenColumnCount").asInt(k.defaultFrozenColumns())
                                : k.defaultFrozenColumns();
                List<String> order = new ArrayList<>();
                if (sn != null && sn.isObject()) {
                    JsonNode arr = sn.get("nonDateColumnOrder");
                    if (arr != null && arr.isArray()) {
                        for (JsonNode item : arr) {
                            String t = item.asText("").trim();
                            if (!t.isEmpty()) {
                                order.add(t);
                            }
                        }
                    }
                }
                sheets.put(k, new SheetPrefs(frozen, order));
            }
            return new ExportPrefs(font, size, theme, sheets);
        } catch (IOException e) {
            return ExportPrefs.defaults();
        }
    }

    public static void save(ExportPrefs prefs) {
        if (prefs == null) {
            return;
        }
        Path store = resolveStore();
        try {
            Files.createDirectories(store.getParent());
            ObjectNode root = JSON.createObjectNode();
            root.put("fontFamily", prefs.fontFamily());
            root.put("fontSizePt", prefs.fontSizePt());
            root.put("themeId", prefs.theme().storedId());
            ObjectNode sheets = root.putObject("sheets");
            for (SheetKey k : SheetKey.values()) {
                SheetPrefs sp = prefs.sheet(k);
                ObjectNode sn = sheets.putObject(k.jsonKey());
                sn.put("frozenColumnCount", sp.frozenColumnCount());
                ArrayNode order = sn.putArray("nonDateColumnOrder");
                for (String t : sp.nonDateColumnOrder()) {
                    order.add(t);
                }
            }
            JSON.writerWithDefaultPrettyPrinter().writeValue(store.toFile(), root);
        } catch (IOException ignored) {
        }
    }

    public static void saveSheetPrefs(SheetKey key, SheetPrefs sheetPrefs) {
        ExportPrefs cur = load();
        Map<SheetKey, SheetPrefs> sheets = new LinkedHashMap<>(cur.sheets());
        sheets.put(key, sheetPrefs != null ? sheetPrefs : SheetPrefs.defaults(key));
        save(
                new ExportPrefs(
                        cur.fontFamily(), cur.fontSizePt(), cur.theme(), sheets));
    }

    public static void resetToDefaults() {
        try {
            Files.deleteIfExists(resolveStore());
        } catch (IOException ignored) {
        }
    }
}
