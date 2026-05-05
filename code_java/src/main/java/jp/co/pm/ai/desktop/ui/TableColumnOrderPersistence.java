package jp.co.pm.ai.desktop.ui;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.function.BooleanSupplier;
import java.util.function.Consumer;
import java.util.function.Supplier;

import javafx.animation.PauseTransition;
import javafx.collections.ListChangeListener;
import javafx.collections.ObservableList;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.util.Duration;

import org.controlsfx.control.spreadsheet.SpreadsheetColumn;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * Persists TableView column order and widths under {@code ~/.pm-ai-desktop/table-column-order.json}.
 *
 * <p>JSON value per key: an array of objects {@code title} + {@code width}. Legacy string-only arrays
 * (column order only) are still read; missing widths use the tab default (e.g. 112px).
 */
public final class TableColumnOrderPersistence {

    private static final ObjectMapper JSON = new ObjectMapper();
    private static final Path STORE =
            Paths.get(System.getProperty("user.home"), ".pm-ai-desktop", "table-column-order.json");

    private static final double MIN_WIDTH = 40.0;

    /** {@link #planResultViewerSheetScopeKey} 用のルート JSON キー接尾辞（見出し固定列数）。 */
    private static final String SCOPE_SUFFIX_HEADER_COLUMN_COUNT = "_headerColumnCount";

    private static final String KEY_PLAN_RESULT_FONT_FAMILY = "planResultViewer_ui_fontFamily";

    private static final String KEY_PLAN_RESULT_FONT_SIZE = "planResultViewer_ui_fontSize";

    private static final String KEY_PLAN_RESULT_ROW_HEIGHT_PCT = "planResultViewer_ui_rowHeightPercent";

    private static final String KEY_PLAN_RESULT_CELL_WRAP = "planResultViewer_ui_cellWrapText";

    /**
     * 計画結果 JSON ビューアの表示設定（フォント・行高・セル折り返し）。
     *
     * @param rowHeightPercent データ行の高さ 50〜2000（100＝既定）
     * @param cellWrapText {@code true} で折り返し、{@code false} で1行・見切れ
     */
    public record PlanResultViewerUiPrefs(
            String family, double size, double rowHeightPercent, boolean cellWrapText) {
        public static PlanResultViewerUiPrefs defaults() {
            return new PlanResultViewerUiPrefs("", 12.0, 100.0, false);
        }
    }

    private static String headerCountKey(TableId id) {
        return id.jsonKey() + "_headerColumnCount";
    }

    public record ColumnSpec(String title, double width) {}

    public enum TableId {
        PLAN_INPUT("planInput"),
        STAGE1_PREVIEW("stage1Preview"),
        ACTUALS_STATUS("actualsStatus"),
        ENV_VARS("envVars"),
        RESULT_DISPATCH_TABLE("resultDispatchTable");

        private final String jsonKey;

        TableId(String jsonKey) {
            this.jsonKey = jsonKey;
        }

        String jsonKey() {
            return jsonKey;
        }
    }

    private TableColumnOrderPersistence() {}

    /**
     * Reorders {@code headersRef} and each row so indices stay aligned (logical column move).
     */
    public static void applyLogicalColumnOrder(
            List<String> headersRef,
            ObservableList<ObservableList<String>> rows,
            List<String> savedOrder) {
        if (savedOrder == null || savedOrder.isEmpty() || headersRef.isEmpty()) {
            return;
        }
        List<String> fileHeaders = new ArrayList<>(headersRef);
        List<Integer> perm = buildPermutation(fileHeaders, savedOrder);
        List<Integer> natural = new ArrayList<>();
        for (int i = 0; i < fileHeaders.size(); i++) {
            natural.add(i);
        }
        if (perm.equals(natural)) {
            return;
        }
        headersRef.clear();
        for (int oldIdx : perm) {
            headersRef.add(fileHeaders.get(oldIdx));
        }
        for (ObservableList<String> row : rows) {
            List<String> old = new ArrayList<>(row);
            row.clear();
            for (int oldIdx : perm) {
                row.add(oldIdx < old.size() ? Objects.requireNonNullElse(old.get(oldIdx), "") : "");
            }
        }
    }

    /**
     * Reorders {@link TableView#getColumns()} to match saved header titles (extras appended).
     */
    @SuppressWarnings("unchecked")
    public static void applyOrderToTableColumns(TableView<?> table, List<String> savedOrder) {
        if (savedOrder == null || savedOrder.isEmpty()) {
            return;
        }
        ObservableList<TableColumn<?, ?>> cols =
                (ObservableList<TableColumn<?, ?>>) (ObservableList<?>) table.getColumns();
        if (cols.isEmpty()) {
            return;
        }
        Map<String, List<TableColumn<?, ?>>> byTitle = new HashMap<>();
        for (TableColumn<?, ?> c : cols) {
            String t = colTitle(c);
            byTitle.computeIfAbsent(t, k -> new ArrayList<>()).add(c);
        }
        List<TableColumn<?, ?>> out = new ArrayList<>();
        Set<TableColumn<?, ?>> used = new HashSet<>();
        for (String title : savedOrder) {
            List<TableColumn<?, ?>> bucket = byTitle.get(title);
            if (bucket != null && !bucket.isEmpty()) {
                TableColumn<?, ?> c = bucket.remove(0);
                if (used.add(c)) {
                    out.add(c);
                }
            }
        }
        for (TableColumn<?, ?> c : cols) {
            if (!used.contains(c)) {
                out.add(c);
            }
        }
        if (out.size() == cols.size()) {
            cols.setAll(out);
        }
    }

    /** Applies {@link ColumnSpec#width()} to columns in current visual order (multiset by title). */
    public static void applyWidthsToTableColumns(
            TableView<?> table, List<ColumnSpec> layout, double defaultWidth) {
        if (layout == null || layout.isEmpty()) {
            return;
        }
        Map<String, ArrayDeque<Double>> q = widthQueues(layout, defaultWidth);
        for (TableColumn<?, ?> c : table.getColumns()) {
            String t = colTitle(c);
            ArrayDeque<Double> dq = q.get(t);
            if (dq != null && !dq.isEmpty()) {
                c.setPrefWidth(dq.poll());
            }
        }
    }

    /** Number of leading visual columns treated as header columns (0 = off). */
    public static int loadHeaderColumnCount(TableId id) {
        try {
            if (!Files.isRegularFile(STORE)) {
                return 0;
            }
            JsonNode root = JSON.readTree(STORE.toFile());
            if (root == null || !root.isObject()) {
                return 0;
            }
            JsonNode n = root.get(headerCountKey(id));
            if (n == null || !n.isNumber()) {
                return 0;
            }
            int v = n.intValue();
            return Math.max(0, Math.min(v, 10_000));
        } catch (IOException e) {
            return 0;
        }
    }

    public static void saveHeaderColumnCount(TableId id, int count) {
        try {
            Files.createDirectories(STORE.getParent());
            ObjectNode root;
            if (Files.isRegularFile(STORE)) {
                JsonNode tree = JSON.readTree(STORE.toFile());
                root =
                        tree != null && tree.isObject()
                                ? (ObjectNode) tree
                                : JSON.createObjectNode();
            } else {
                root = JSON.createObjectNode();
            }
            int v = Math.max(0, Math.min(count, 10_000));
            root.put(headerCountKey(id), v);
            JSON.writerWithDefaultPrettyPrinter().writeValue(STORE.toFile(), root);
        } catch (IOException ignored) {
        }
    }

    public static List<ColumnSpec> loadLayout(TableId id) {
        try {
            if (!Files.isRegularFile(STORE)) {
                return List.of();
            }
            JsonNode root = JSON.readTree(STORE.toFile());
            if (root == null || !root.isObject()) {
                return List.of();
            }
            JsonNode arr = root.get(id.jsonKey());
            return parseLayoutArray(arr, defaultWidthFallback());
        } catch (IOException e) {
            return List.of();
        }
    }

    /**
     * 計画結果ビューアのシート単位スコープキー（データセット種別 + シート名）。列順・列幅・見出し列数の保存に使う。
     */
    public static String planResultViewerSheetScopeKey(String datasetTag, String sheetName) {
        return "prvSh_"
                + sanitizeScopeFragment(datasetTag)
                + "_"
                + sanitizeScopeFragment(sheetName);
    }

    private static String sanitizeScopeFragment(String s) {
        if (s == null) {
            return "null";
        }
        String t = s.strip();
        if (t.length() > 120) {
            t = t.substring(0, 120);
        }
        StringBuilder sb = new StringBuilder(Math.max(8, t.length()));
        for (int i = 0; i < t.length(); i++) {
            char ch = t.charAt(i);
            if (ch == '"' || ch == '\\') {
                sb.append('_');
            } else if (ch < 32) {
                sb.append('_');
            } else {
                sb.append(ch);
            }
        }
        return sb.length() > 0 ? sb.toString() : "x";
    }

    /** {@link #planResultViewerSheetScopeKey} に対応する列レイアウト（順・幅）。 */
    public static List<ColumnSpec> loadLayoutForScope(String scopeKey) {
        try {
            if (!Files.isRegularFile(STORE)) {
                return List.of();
            }
            JsonNode root = JSON.readTree(STORE.toFile());
            if (root == null || !root.isObject()) {
                return List.of();
            }
            JsonNode arr = root.get(scopeKey);
            return parseLayoutArray(arr, defaultWidthFallback());
        } catch (IOException e) {
            return List.of();
        }
    }

    /** {@link #planResultViewerSheetScopeKey} に列レイアウトを保存する（他キーは維持）。 */
    public static void saveLayoutForScope(String scopeKey, List<ColumnSpec> columns) {
        try {
            Files.createDirectories(STORE.getParent());
            ObjectNode root;
            if (Files.isRegularFile(STORE)) {
                JsonNode tree = JSON.readTree(STORE.toFile());
                root =
                        tree != null && tree.isObject()
                                ? (ObjectNode) tree.deepCopy()
                                : JSON.createObjectNode();
            } else {
                root = JSON.createObjectNode();
            }
            ArrayNode arr = JSON.createArrayNode();
            double def = defaultWidthFallback();
            for (ColumnSpec c : columns) {
                ObjectNode o = JSON.createObjectNode();
                o.put("title", c.title());
                o.put("width", clampWidth(c.width(), def));
                arr.add(o);
            }
            root.set(scopeKey, arr);
            JSON.writerWithDefaultPrettyPrinter().writeValue(STORE.toFile(), root);
        } catch (IOException ignored) {
        }
    }

    public static int loadHeaderColumnCountForScope(String sheetScopeKey) {
        try {
            if (!Files.isRegularFile(STORE)) {
                return 0;
            }
            JsonNode root = JSON.readTree(STORE.toFile());
            if (root == null || !root.isObject()) {
                return 0;
            }
            JsonNode n = root.get(sheetScopeKey + SCOPE_SUFFIX_HEADER_COLUMN_COUNT);
            if (n == null || !n.isNumber()) {
                return 0;
            }
            int v = n.intValue();
            return Math.max(0, Math.min(v, 10_000));
        } catch (IOException e) {
            return 0;
        }
    }

    public static void saveHeaderColumnCountForScope(String sheetScopeKey, int count) {
        try {
            Files.createDirectories(STORE.getParent());
            ObjectNode root;
            if (Files.isRegularFile(STORE)) {
                JsonNode tree = JSON.readTree(STORE.toFile());
                root =
                        tree != null && tree.isObject()
                                ? (ObjectNode) tree.deepCopy()
                                : JSON.createObjectNode();
            } else {
                root = JSON.createObjectNode();
            }
            int v = Math.max(0, Math.min(count, 10_000));
            root.put(sheetScopeKey + SCOPE_SUFFIX_HEADER_COLUMN_COUNT, v);
            JSON.writerWithDefaultPrettyPrinter().writeValue(STORE.toFile(), root);
        } catch (IOException ignored) {
        }
    }

    public static PlanResultViewerUiPrefs loadPlanResultViewerUiPrefs() {
        try {
            if (!Files.isRegularFile(STORE)) {
                return PlanResultViewerUiPrefs.defaults();
            }
            JsonNode root = JSON.readTree(STORE.toFile());
            if (root == null || !root.isObject()) {
                return PlanResultViewerUiPrefs.defaults();
            }
            String fam = root.path(KEY_PLAN_RESULT_FONT_FAMILY).asText("");
            double sz = root.path(KEY_PLAN_RESULT_FONT_SIZE).asDouble(12.0);
            if (Double.isNaN(sz) || sz < 6) {
                sz = 12.0;
            }
            if (sz > 96) {
                sz = 96;
            }
            double rh = root.path(KEY_PLAN_RESULT_ROW_HEIGHT_PCT).asDouble(100.0);
            if (Double.isNaN(rh) || rh < 50) {
                rh = 100.0;
            }
            if (rh > SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MAX) {
                rh = SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MAX;
            }
            boolean wrap = root.path(KEY_PLAN_RESULT_CELL_WRAP).asBoolean(false);
            return new PlanResultViewerUiPrefs(fam != null ? fam : "", sz, rh, wrap);
        } catch (IOException e) {
            return PlanResultViewerUiPrefs.defaults();
        }
    }

    public static void savePlanResultViewerUiPrefs(PlanResultViewerUiPrefs prefs) {
        if (prefs == null) {
            return;
        }
        try {
            Files.createDirectories(STORE.getParent());
            ObjectNode root;
            if (Files.isRegularFile(STORE)) {
                JsonNode tree = JSON.readTree(STORE.toFile());
                root =
                        tree != null && tree.isObject()
                                ? (ObjectNode) tree.deepCopy()
                                : JSON.createObjectNode();
            } else {
                root = JSON.createObjectNode();
            }
            root.put(KEY_PLAN_RESULT_FONT_FAMILY, prefs.family() != null ? prefs.family() : "");
            root.put(KEY_PLAN_RESULT_FONT_SIZE, prefs.size());
            double rh = prefs.rowHeightPercent();
            if (Double.isNaN(rh)) {
                rh = 100.0;
            }
            rh =
                    Math.min(
                            SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MAX,
                            Math.max(SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MIN, rh));
            root.put(KEY_PLAN_RESULT_ROW_HEIGHT_PCT, rh);
            root.put(KEY_PLAN_RESULT_CELL_WRAP, prefs.cellWrapText());
            JSON.writerWithDefaultPrettyPrinter().writeValue(STORE.toFile(), root);
        } catch (IOException ignored) {
        }
    }

    public static List<Double> resolveWidthsForHeaders(
            List<String> headersInOrder, List<ColumnSpec> layout, double defaultWidth) {
        double def = normalizeDefaultWidth(defaultWidth);
        if (layout == null || layout.isEmpty()) {
            return headersInOrder.stream().map(h -> def).toList();
        }
        Map<String, ArrayDeque<Double>> q = widthQueues(layout, def);
        List<Double> out = new ArrayList<>();
        for (String h : headersInOrder) {
            ArrayDeque<Double> dq = q.get(h);
            if (dq != null && !dq.isEmpty()) {
                out.add(dq.poll());
            } else {
                out.add(def);
            }
        }
        return out;
    }

    public static void saveLayout(TableId id, List<ColumnSpec> columns) {
        try {
            Files.createDirectories(STORE.getParent());
            ObjectNode root;
            if (Files.isRegularFile(STORE)) {
                JsonNode tree = JSON.readTree(STORE.toFile());
                root =
                        tree != null && tree.isObject()
                                ? (ObjectNode) tree.deepCopy()
                                : JSON.createObjectNode();
            } else {
                root = JSON.createObjectNode();
            }
            ArrayNode arr = JSON.createArrayNode();
            double def = defaultWidthFallback();
            for (ColumnSpec c : columns) {
                ObjectNode o = JSON.createObjectNode();
                o.put("title", c.title());
                o.put("width", clampWidth(c.width(), def));
                arr.add(o);
            }
            root.set(id.jsonKey(), arr);
            JSON.writerWithDefaultPrettyPrinter().writeValue(STORE.toFile(), root);
        } catch (IOException ignored) {
        }
    }

    public static List<ColumnSpec> snapshot(TableView<?> table) {
        List<ColumnSpec> out = new ArrayList<>();
        for (TableColumn<?, ?> c : table.getColumns()) {
            out.add(new ColumnSpec(colTitle(c), effectiveWidth(c)));
        }
        return out;
    }

    /** Column widths for ControlsFX {@link SpreadsheetView} (titles from sheet header list). */
    public static List<ColumnSpec> snapshotSpreadsheet(SpreadsheetView view, List<String> headerTitlesInOrder) {
        List<ColumnSpec> out = new ArrayList<>();
        if (view == null || headerTitlesInOrder == null) {
            return out;
        }
        ObservableList<SpreadsheetColumn> cols = view.getColumns();
        for (int i = 0; i < cols.size(); i++) {
            String title = i < headerTitlesInOrder.size() ? headerTitlesInOrder.get(i) : "";
            double w = cols.get(i).getWidth();
            if (w <= 1 || Double.isNaN(w)) {
                w = defaultWidthFallback();
            }
            out.add(new ColumnSpec(title, clampWidth(w, defaultWidthFallback())));
        }
        return out;
    }

    /**
     * Saves spreadsheet column widths (order follows {@code headerTitlesSupplier} / grid columns).
     */
    public static void installSpreadsheetColumnLayoutWatcher(
            SpreadsheetView view,
            TableId id,
            BooleanSupplier suppressSave,
            Supplier<List<String>> headerTitlesSupplier) {
        installSpreadsheetColumnLayoutWatcherImpl(
                view,
                suppressSave,
                headerTitlesSupplier,
                layout -> saveLayout(id, layout));
    }

    /**
     * 任意ストレージキー（例: 計画結果ビューアのシート単位 {@link #planResultViewerSheetScopeKey}）へ列幅を保存する。
     */
    public static void installSpreadsheetColumnLayoutWatcherForScope(
            SpreadsheetView view,
            String scopeStorageKey,
            BooleanSupplier suppressSave,
            Supplier<List<String>> headerTitlesSupplier) {
        installSpreadsheetColumnLayoutWatcherImpl(
                view,
                suppressSave,
                headerTitlesSupplier,
                layout -> saveLayoutForScope(scopeStorageKey, layout));
    }

    private static void installSpreadsheetColumnLayoutWatcherImpl(
            SpreadsheetView view,
            BooleanSupplier suppressSave,
            Supplier<List<String>> headerTitlesSupplier,
            Consumer<List<ColumnSpec>> layoutSaver) {
        PauseTransition debounce = new PauseTransition(Duration.millis(400));
        Runnable flushWidths =
                () -> {
                    if (suppressSave.getAsBoolean()) {
                        return;
                    }
                    layoutSaver.accept(snapshotSpreadsheet(view, headerTitlesSupplier.get()));
                };
        debounce.setOnFinished(e -> flushWidths.run());

        Runnable scheduleWidthSave =
                () -> {
                    if (suppressSave.getAsBoolean()) {
                        return;
                    }
                    debounce.stop();
                    debounce.playFromStart();
                };

        view.getColumns()
                .addListener(
                        (ListChangeListener<SpreadsheetColumn>)
                                c -> {
                                    if (suppressSave.getAsBoolean()) {
                                        return;
                                    }
                                    while (c.next()) {
                                        if (c.wasAdded()) {
                                            for (SpreadsheetColumn col : c.getAddedSubList()) {
                                                attachSpreadsheetWidthDebounced(
                                                        col, scheduleWidthSave, suppressSave);
                                            }
                                        }
                                        if (c.wasPermutated() || c.wasAdded() || c.wasRemoved()) {
                                            debounce.stop();
                                            flushWidths.run();
                                            return;
                                        }
                                    }
                                });

        for (SpreadsheetColumn col : view.getColumns()) {
            attachSpreadsheetWidthDebounced(col, scheduleWidthSave, suppressSave);
        }
    }

    private static void attachSpreadsheetWidthDebounced(
            SpreadsheetColumn col, Runnable scheduleWidthSave, BooleanSupplier suppressSave) {
        col.widthProperty()
                .addListener(
                        (obs, o, n) -> {
                            if (suppressSave.getAsBoolean()) {
                                return;
                            }
                            scheduleWidthSave.run();
                        });
    }

    /**
     * Saves column order and widths. Structural changes save immediately; width-only changes are debounced.
     */
    public static void installColumnLayoutWatcher(
            TableView<?> table, TableId id, BooleanSupplier suppressSave) {
        PauseTransition debounce = new PauseTransition(Duration.millis(400));
        Runnable flushWidths =
                () -> {
                    if (suppressSave.getAsBoolean()) {
                        return;
                    }
                    saveLayout(id, snapshot(table));
                };
        debounce.setOnFinished(e -> flushWidths.run());

        Runnable scheduleWidthSave =
                () -> {
                    if (suppressSave.getAsBoolean()) {
                        return;
                    }
                    debounce.stop();
                    debounce.playFromStart();
                };

        table.getColumns()
                .addListener(
                        (ListChangeListener<TableColumn<?, ?>>)
                                c -> {
                                    if (suppressSave.getAsBoolean()) {
                                        return;
                                    }
                                    while (c.next()) {
                                        if (c.wasAdded()) {
                                            for (TableColumn<?, ?> col : c.getAddedSubList()) {
                                                attachWidthDebounced(col, scheduleWidthSave, suppressSave);
                                            }
                                        }
                                        if (c.wasPermutated() || c.wasAdded() || c.wasRemoved()) {
                                            debounce.stop();
                                            flushWidths.run();
                                            return;
                                        }
                                    }
                                });

        for (TableColumn<?, ?> col : table.getColumns()) {
            attachWidthDebounced(col, scheduleWidthSave, suppressSave);
        }
    }

    private static void attachWidthDebounced(
            TableColumn<?, ?> col, Runnable scheduleWidthSave, BooleanSupplier suppressSave) {
        col.widthProperty()
                .addListener(
                        (obs, o, n) -> {
                            if (suppressSave.getAsBoolean()) {
                                return;
                            }
                            scheduleWidthSave.run();
                        });
    }

    private static List<ColumnSpec> parseLayoutArray(JsonNode arr, double defaultW) {
        if (arr == null || !arr.isArray() || arr.isEmpty()) {
            return List.of();
        }
        JsonNode first = arr.get(0);
        if (first != null && first.isTextual()) {
            List<ColumnSpec> out = new ArrayList<>();
            for (JsonNode el : arr) {
                if (el != null && el.isTextual()) {
                    out.add(new ColumnSpec(el.asText(""), defaultW));
                }
            }
            return out;
        }
        List<ColumnSpec> out = new ArrayList<>();
        for (JsonNode el : arr) {
            if (el != null && el.isObject()) {
                String t = el.path("title").asText("");
                double w = el.path("width").asDouble(-1);
                out.add(new ColumnSpec(t, w > 0 ? w : defaultW));
            }
        }
        return out;
    }

    private static Map<String, ArrayDeque<Double>> widthQueues(List<ColumnSpec> layout, double defaultWidth) {
        double def = normalizeDefaultWidth(defaultWidth);
        Map<String, ArrayDeque<Double>> q = new HashMap<>();
        for (ColumnSpec c : layout) {
            double w = c.width() > 0 && !Double.isNaN(c.width()) ? clampWidth(c.width(), def) : def;
            q.computeIfAbsent(c.title(), t -> new ArrayDeque<>()).add(w);
        }
        return q;
    }

    private static double effectiveWidth(TableColumn<?, ?> c) {
        double w = c.getWidth();
        if (w > 1 && !Double.isNaN(w)) {
            return clampWidth(w, defaultWidthFallback());
        }
        w = c.getPrefWidth();
        if (w > 1 && !Double.isNaN(w)) {
            return clampWidth(w, defaultWidthFallback());
        }
        return defaultWidthFallback();
    }

    private static double clampWidth(double w, double defaultW) {
        if (w <= 0 || Double.isNaN(w)) {
            return Math.max(MIN_WIDTH, defaultW);
        }
        return Math.max(MIN_WIDTH, w);
    }

    private static double normalizeDefaultWidth(double defaultWidth) {
        if (defaultWidth <= 0 || Double.isNaN(defaultWidth)) {
            return 112.0;
        }
        return Math.max(MIN_WIDTH, defaultWidth);
    }

    private static double defaultWidthFallback() {
        return 112.0;
    }

    private static List<Integer> buildPermutation(List<String> fileHeaders, List<String> savedOrder) {
        List<Integer> perm = new ArrayList<>();
        Set<Integer> used = new HashSet<>();
        for (String title : savedOrder) {
            int idx = findNextUnusedMatching(fileHeaders, title, used);
            if (idx >= 0) {
                perm.add(idx);
                used.add(idx);
            }
        }
        for (int i = 0; i < fileHeaders.size(); i++) {
            if (!used.contains(i)) {
                perm.add(i);
                used.add(i);
            }
        }
        return perm;
    }

    private static int findNextUnusedMatching(List<String> headers, String title, Set<Integer> used) {
        for (int i = 0; i < headers.size(); i++) {
            if (used.contains(i)) {
                continue;
            }
            if (Objects.equals(headers.get(i), title)) {
                return i;
            }
        }
        return -1;
    }

    private static String colTitle(TableColumn<?, ?> c) {
        String t = c.getText();
        return t != null ? t : "";
    }
}
