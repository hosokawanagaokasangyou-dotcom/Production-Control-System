package jp.co.pm.ai.desktop.ui;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.function.BooleanSupplier;
import java.util.function.Consumer;
import java.util.function.Supplier;
import java.util.Iterator;

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

import jp.co.pm.ai.desktop.config.InitSettingPaths;

/**
 * Persists TableView column order and widths under {@code ~/.pm-ai-desktop/table-column-order.json}.
 *
 * <p>初回インストールでは {@code pm-ai-data/config/bundled_table_column_order.json}（またはクラスパス上の同名）を
 * {@link #materializeBundledDefaultsIfStoreMissing()} でユーザーホームへコピーしてから各タブが読み込む。
 *
 * <p>ポータブル自動バージョンアップ後は {@link #overwriteTableColumnOrderStoreFromBundledAfterPortableUpgrade()} で
 * {@link #STORE} を正本バンドルに合わせて上書きする。
 *
 * <p>JSON value per key: an array of objects {@code title} + {@code width}. Legacy string-only arrays
 * (column order only) are still read; missing widths use the tab default (e.g. 112px).
 */
public final class TableColumnOrderPersistence {

    private static final ObjectMapper JSON = new ObjectMapper();
    private static final Path STORE =
            Paths.get(System.getProperty("user.home"), ".pm-ai-desktop", "table-column-order.json");

    /** 現在の列順ストアを読む（無ければ空オブジェクト）。ユーザープロファイル保存用。 */
    public static JsonNode readCurrentStoreRoot() {
        try {
            if (!Files.isRegularFile(STORE)) {
                return JSON.createObjectNode();
            }
            JsonNode n = JSON.readTree(STORE.toFile());
            return n != null && n.isObject() ? n : JSON.createObjectNode();
        } catch (IOException e) {
            return JSON.createObjectNode();
        }
    }

    /** 列順ストアを丸ごと置換する（ユーザープロファイル読み出し用）。 */
    public static void overwriteStoreRoot(JsonNode root) throws IOException {
        ObjectNode out =
                root != null && root.isObject()
                        ? (ObjectNode) root.deepCopy()
                        : JSON.createObjectNode();
        Files.createDirectories(STORE.getParent());
        JSON.writerWithDefaultPrettyPrinter().writeValue(STORE.toFile(), out);
    }

    private static final String BUNDLED_TABLE_COLUMN_ORDER_RESOURCE =
            "/jp/co/pm/ai/desktop/config/bundled_table_column_order.json";

    private static final double MIN_WIDTH = 40.0;

    /** {@link #planResultViewerSheetScopeKey} 用のルート JSON キー接尾辞（見出し固定列数）。 */
    private static final String SCOPE_SUFFIX_HEADER_COLUMN_COUNT = "_headerColumnCount";

    /** {@link #planResultViewerSheetScopeKey} 用の列フィルタ状態（許容値）。 */
    private static final String SCOPE_SUFFIX_COLUMN_FILTERS = "_columnFilters";

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

    /**
     * 計画結果 JSON ビューアの列フィルタ（見出し名と許容セル文字列）。並べ替え状態は含まない。
     *
     * @param title 列見出し（論理列の識別）
     * @param values 表示を許可するセル文字列の集合をリスト化したもの
     */
    public record PlanResultViewerColumnFilterSpec(String title, List<String> values) {}

    /**
     * 段階1成形結果・結果_配台表 JSON など、ControlsFX Spreadsheet タブのデータ行の高さと折り返し。
     *
     * @param rowHeightPercent {@link SpreadsheetTabularSupport#PLAN_RESULT_ROW_HEIGHT_PCT_MIN} 〜 {@link
     *     SpreadsheetTabularSupport#PLAN_RESULT_ROW_HEIGHT_PCT_MAX}（100＝既定）
     */
    public record SpreadsheetTabPresentationPrefs(double rowHeightPercent, boolean cellWrapText) {
        public static SpreadsheetTabPresentationPrefs defaults() {
            return new SpreadsheetTabPresentationPrefs(100.0, false);
        }
    }

    private static String spreadsheetTabRowHeightKey(TableId id) {
        return id.jsonKey() + "_ui_rowHeightPercent";
    }

    private static String spreadsheetTabCellWrapKey(TableId id) {
        return id.jsonKey() + "_ui_cellWrapText";
    }

    /**
     * JSON suffix unchanged ({@code _ui_productConditionBreakdownFilter}) for backward compatibility with existing
     * {@link #STORE} files; column meaning is 「製造条件(内訳)」.
     */
    private static String processingActualsManufacturingConditionBreakdownFilterStoreKey(TableId id) {
        return id.jsonKey() + "_ui_productConditionBreakdownFilter";
    }

    /**
     * 加工実績タブ「製造条件(内訳)」コンボの選択値（{@link TableId#PROCESSING_ACTUALS_DETAIL_RAW}）。空は「（全行）」。
     */
    public static String loadProcessingActualsManufacturingConditionBreakdownFilter() {
        TableId id = TableId.PROCESSING_ACTUALS_DETAIL_RAW;
        try {
            if (!Files.isRegularFile(STORE)) {
                return "";
            }
            JsonNode root = JSON.readTree(STORE.toFile());
            if (root == null || !root.isObject()) {
                return "";
            }
            return root.path(processingActualsManufacturingConditionBreakdownFilterStoreKey(id)).asText("");
        } catch (IOException e) {
            return "";
        }
    }

    /** Persists {@link #loadProcessingActualsManufacturingConditionBreakdownFilter()} into {@link #STORE}. */
    public static void saveProcessingActualsManufacturingConditionBreakdownFilter(String text) {
        TableId id = TableId.PROCESSING_ACTUALS_DETAIL_RAW;
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
            root.put(
                    processingActualsManufacturingConditionBreakdownFilterStoreKey(id),
                    text != null ? text : "");
            JSON.writerWithDefaultPrettyPrinter().writeValue(STORE.toFile(), root);
        } catch (IOException ignored) {
        }
    }

    /** 指定 {@link TableId} の Spreadsheet 行高・折り返しを読み込む（キーが無ければ既定）。 */
    public static SpreadsheetTabPresentationPrefs loadSpreadsheetTabPresentationPrefs(TableId id) {
        if (id == null) {
            return SpreadsheetTabPresentationPrefs.defaults();
        }
        try {
            if (!Files.isRegularFile(STORE)) {
                return SpreadsheetTabPresentationPrefs.defaults();
            }
            JsonNode root = JSON.readTree(STORE.toFile());
            if (root == null || !root.isObject()) {
                return SpreadsheetTabPresentationPrefs.defaults();
            }
            double rh = root.path(spreadsheetTabRowHeightKey(id)).asDouble(100.0);
            if (Double.isNaN(rh) || rh < SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MIN) {
                rh = 100.0;
            }
            rh =
                    Math.min(
                            SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MAX,
                            Math.max(SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MIN, rh));
            boolean wrap = root.path(spreadsheetTabCellWrapKey(id)).asBoolean(false);
            return new SpreadsheetTabPresentationPrefs(rh, wrap);
        } catch (IOException e) {
            return SpreadsheetTabPresentationPrefs.defaults();
        }
    }

    /**
     * 納期管理ビュー「アラ・実績・シス比較」の日付ウィンドウ（本日基準の過去日数・未来日数）。
     *
     * <p>環境変数 {@code PM_AI_DELIVERY_CALENDAR_PAST_DAYS} / {@code PM_AI_DELIVERY_CALENDAR_FUTURE_DAYS}
     * の上書き値として {@code pm_ai_delivery_calendar_view.py} に渡される。
     */
    public record DeliveryCalendarDateWindowPrefs(int pastDays, int futureDays) {
        public static final int MIN = 1;
        public static final int MAX = 800;
        public static final int DEFAULT_PAST = 14;
        public static final int DEFAULT_FUTURE = 30;

        public static DeliveryCalendarDateWindowPrefs defaults() {
            return new DeliveryCalendarDateWindowPrefs(DEFAULT_PAST, DEFAULT_FUTURE);
        }

        public DeliveryCalendarDateWindowPrefs {
            pastDays = Math.max(MIN, Math.min(MAX, pastDays));
            futureDays = Math.max(MIN, Math.min(MAX, futureDays));
        }
    }

    private static final String DELIVERY_CALENDAR_PAST_DAYS_KEY =
            "deliveryCalendarMain_ui_pastDays";
    private static final String DELIVERY_CALENDAR_FUTURE_DAYS_KEY =
            "deliveryCalendarMain_ui_futureDays";
    private static final String DELIVERY_CALENDAR_FONT_FAMILY_KEY =
            "deliveryCalendarMain_ui_fontFamily";

    /**
     * 納期管理ビュー「アラ・実績・シス比較」表のフォントファミリ名（空文字＝既定）。
     * 値変更時に SpreadsheetView 全体へインライン {@code -fx-font-family} で適用される。
     */
    public static String loadDeliveryCalendarFontFamily() {
        try {
            if (!Files.isRegularFile(STORE)) {
                return "";
            }
            JsonNode root = JSON.readTree(STORE.toFile());
            if (root == null || !root.isObject()) {
                return "";
            }
            return root.path(DELIVERY_CALENDAR_FONT_FAMILY_KEY).asText("");
        } catch (IOException e) {
            return "";
        }
    }

    public static void saveDeliveryCalendarFontFamily(String family) {
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
            root.put(DELIVERY_CALENDAR_FONT_FAMILY_KEY, family != null ? family : "");
            JSON.writerWithDefaultPrettyPrinter().writeValue(STORE.toFile(), root);
        } catch (IOException ignored) {
        }
    }

    /** 納期管理ビュー「アラ・実績・シス比較」の日付ウィンドウを読み込む（キーが無ければ既定 14/30）。 */
    public static DeliveryCalendarDateWindowPrefs loadDeliveryCalendarDateWindowPrefs() {
        try {
            if (!Files.isRegularFile(STORE)) {
                return DeliveryCalendarDateWindowPrefs.defaults();
            }
            JsonNode root = JSON.readTree(STORE.toFile());
            if (root == null || !root.isObject()) {
                return DeliveryCalendarDateWindowPrefs.defaults();
            }
            int past =
                    root.path(DELIVERY_CALENDAR_PAST_DAYS_KEY)
                            .asInt(DeliveryCalendarDateWindowPrefs.DEFAULT_PAST);
            int future =
                    root.path(DELIVERY_CALENDAR_FUTURE_DAYS_KEY)
                            .asInt(DeliveryCalendarDateWindowPrefs.DEFAULT_FUTURE);
            return new DeliveryCalendarDateWindowPrefs(past, future);
        } catch (IOException e) {
            return DeliveryCalendarDateWindowPrefs.defaults();
        }
    }

    /** 納期管理ビュー「アラ・実績・シス比較」の日付ウィンドウを永続化する。 */
    public static void saveDeliveryCalendarDateWindowPrefs(DeliveryCalendarDateWindowPrefs prefs) {
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
            root.put(DELIVERY_CALENDAR_PAST_DAYS_KEY, prefs.pastDays());
            root.put(DELIVERY_CALENDAR_FUTURE_DAYS_KEY, prefs.futureDays());
            JSON.writerWithDefaultPrettyPrinter().writeValue(STORE.toFile(), root);
        } catch (IOException ignored) {
        }
    }

    /** 指定 {@link TableId} の Spreadsheet 行高・折り返しを永続化する。 */
    public static void saveSpreadsheetTabPresentationPrefs(
            TableId id, SpreadsheetTabPresentationPrefs prefs) {
        if (id == null || prefs == null) {
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
            double rh = prefs.rowHeightPercent();
            if (Double.isNaN(rh)) {
                rh = 100.0;
            }
            rh =
                    Math.min(
                            SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MAX,
                            Math.max(SpreadsheetTabularSupport.PLAN_RESULT_ROW_HEIGHT_PCT_MIN, rh));
            root.put(spreadsheetTabRowHeightKey(id), rh);
            root.put(spreadsheetTabCellWrapKey(id), prefs.cellWrapText());
            JSON.writerWithDefaultPrettyPrinter().writeValue(STORE.toFile(), root);
        } catch (IOException ignored) {
        }
    }

    private static String headerCountKey(TableId id) {
        return id.jsonKey() + "_headerColumnCount";
    }

    private static String columnVisibleKey(TableId id) {
        return id.jsonKey() + "_columnVisible";
    }

    /** {@link #planResultViewerSheetScopeKey} 単位の列ごとの表示フラグ（boolean[] と JSON 配列）。 */
    private static final String SCOPE_SUFFIX_COLUMN_VISIBLE = "_columnVisible";

    /**
     * 列の表示状態（{@code headersRef} の並びと同一長）。既定はすべて {@code true}。
     */
    public static boolean[] loadColumnVisibility(TableId id, int columnCount) {
        boolean[] defaults = visibilityDefaults(columnCount);
        if (id == null || columnCount <= 0) {
            return defaults;
        }
        try {
            if (!Files.isRegularFile(STORE)) {
                return defaults;
            }
            JsonNode root = JSON.readTree(STORE.toFile());
            if (root == null || !root.isObject()) {
                return defaults;
            }
            JsonNode arr = root.get(columnVisibleKey(id));
            return mergeVisibilityFromJsonArray(defaults, arr);
        } catch (IOException e) {
            return defaults;
        }
    }

    public static void saveColumnVisibility(TableId id, boolean[] visible) {
        if (id == null || visible == null) {
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
            ArrayNode arr = JSON.createArrayNode();
            for (boolean b : visible) {
                arr.add(b);
            }
            root.set(columnVisibleKey(id), arr);
            JSON.writerWithDefaultPrettyPrinter().writeValue(STORE.toFile(), root);
        } catch (IOException ignored) {
        }
    }

    public static boolean[] loadColumnVisibilityForScope(String scopeKey, int columnCount) {
        boolean[] defaults = visibilityDefaults(columnCount);
        if (scopeKey == null || scopeKey.isEmpty() || columnCount <= 0) {
            return defaults;
        }
        try {
            if (!Files.isRegularFile(STORE)) {
                return defaults;
            }
            JsonNode root = JSON.readTree(STORE.toFile());
            if (root == null || !root.isObject()) {
                return defaults;
            }
            JsonNode arr = root.get(scopeKey + SCOPE_SUFFIX_COLUMN_VISIBLE);
            return mergeVisibilityFromJsonArray(defaults, arr);
        } catch (IOException e) {
            return defaults;
        }
    }

    public static void saveColumnVisibilityForScope(String scopeKey, boolean[] visible) {
        if (scopeKey == null || scopeKey.isEmpty() || visible == null) {
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
            ArrayNode arr = JSON.createArrayNode();
            for (boolean b : visible) {
                arr.add(b);
            }
            root.set(scopeKey + SCOPE_SUFFIX_COLUMN_VISIBLE, arr);
            JSON.writerWithDefaultPrettyPrinter().writeValue(STORE.toFile(), root);
        } catch (IOException ignored) {
        }
    }

    private static boolean[] visibilityDefaults(int columnCount) {
        if (columnCount <= 0) {
            return new boolean[0];
        }
        boolean[] d = new boolean[columnCount];
        Arrays.fill(d, true);
        return d;
    }

    private static boolean[] mergeVisibilityFromJsonArray(boolean[] defaults, JsonNode arr) {
        if (arr == null || !arr.isArray()) {
            return defaults;
        }
        for (int i = 0; i < defaults.length && i < arr.size(); i++) {
            defaults[i] = arr.get(i).asBoolean(true);
        }
        return defaults;
    }

    /**
     * {@link #applyLogicalColumnOrder} と同じ並べ替えで {@code visibilityBefore} を並べ替える。
     *
     * @param headersBefore {@code applyLogicalColumnOrder} 直前の {@code headersRef} のコピー
     * @param visibilityBefore {@code headersBefore} と同一長（短い場合は true で埋める）
     * @param savedOrderTitles {@code applyLogicalColumnOrder} に渡す順序
     */
    public static boolean[] permuteVisibilityForLogicalReorder(
            List<String> headersBefore,
            boolean[] visibilityBefore,
            List<String> savedOrderTitles) {
        if (headersBefore == null || savedOrderTitles == null) {
            return visibilityBefore;
        }
        int n = headersBefore.size();
        boolean[] src = visibilityDefaults(n);
        if (visibilityBefore != null) {
            for (int i = 0; i < n; i++) {
                src[i] = i < visibilityBefore.length ? visibilityBefore[i] : true;
            }
        }
        List<Integer> perm = buildPermutation(new ArrayList<>(headersBefore), savedOrderTitles);
        if (perm.size() != n) {
            return src;
        }
        boolean[] out = new boolean[n];
        for (int j = 0; j < perm.size(); j++) {
            int oldIdx = perm.get(j);
            out[j] = src[oldIdx];
        }
        return out;
    }

    public record ColumnSpec(String title, double width) {}

    public enum TableId {
        PLAN_INPUT("planInput"),
        STAGE1_PREVIEW("stage1Preview"),
        ACTUALS_STATUS("actualsStatus"),
        ENV_VARS("envVars"),
        RESULT_DISPATCH_TABLE("resultDispatchTable"),
        /** PM_AI_TASK_INPUT_SOURCE_DIR 最新ファイルの生表（アラジン加工計画取得データ） */
        ALADDIN_PROCESSING_PLAN_RAW("aladdinProcessingPlanRaw"),
        /** PM_AI_ACTUAL_DETAIL_* で解決した加工実績明細ブックの生表（納期管理ビュー子タブ） */
        PROCESSING_ACTUALS_DETAIL_RAW("processingActualsDetailRaw"),
        /** 配台計画手動修正タブ「タスク×日付」 */
        DISPATCH_INTERACTIVE_WIDE("dispatchInteractiveWide"),
        /** 配台計画手動修正タブ「工程+機械×日」 */
        DISPATCH_INTERACTIVE_BY_DAY("dispatchInteractiveByDay"),
        /** 納期管理ビュー「カレンダー」メイン表 */
        DELIVERY_CALENDAR_MAIN("deliveryCalendarMain"),
        /** 納期管理ビュー「計画比較」別表 */
        DELIVERY_CALENDAR_COMPARE("deliveryCalendarCompare");

        private final String jsonKey;

        TableId(String jsonKey) {
            this.jsonKey = jsonKey;
        }

        String jsonKey() {
            return jsonKey;
        }
    }

    private TableColumnOrderPersistence() {}

    /** {@link #STORE} — export / factory reset で削除するときに参照。 */
    public static Path userHomeStorePath() {
        return STORE;
    }

    /**
     * init_setting 書き出し用: バンドル既定（クラスパス・{@code pm-ai-data/config}・既存 {@code init_setting}）
     * に、ユーザーホームの {@link #STORE} を上乗せマージしたルート（ユーザー値が優先）。
     */
    public static JsonNode mergedTableColumnDefaultsRootForExport() {
        ObjectNode acc = JSON.createObjectNode();
        JsonNode bundled = readBundledTableColumnOrderRoot();
        if (bundled != null && bundled.isObject()) {
            deepMergeTableColumnRoot(acc, (ObjectNode) bundled);
        }
        try {
            if (Files.isRegularFile(STORE)) {
                JsonNode user = JSON.readTree(STORE.toFile());
                if (user != null && user.isObject()) {
                    deepMergeTableColumnRoot(acc, (ObjectNode) user);
                }
            }
        } catch (IOException ignored) {
        }
        return acc;
    }

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
     * Same as {@link #applyLogicalColumnOrder} for delivery-calendar main-grid rows ({@link
     * DeliveryCalendarMainCell}).
     */
    public static void applyLogicalColumnOrderDeliveryCalendar(
            List<String> headersRef,
            ObservableList<ObservableList<DeliveryCalendarMainCell>> rows,
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
        DeliveryCalendarMainCell emptyCell = new DeliveryCalendarMainCell.PlainText("");
        for (ObservableList<DeliveryCalendarMainCell> row : rows) {
            List<DeliveryCalendarMainCell> old = new ArrayList<>(row);
            row.clear();
            for (int oldIdx : perm) {
                row.add(
                        oldIdx < old.size()
                                ? Objects.requireNonNullElse(old.get(oldIdx), emptyCell)
                                : emptyCell);
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

    /** {@link #planResultViewerSheetScopeKey} 単位で保存した列フィルタを読む。 */
    public static List<PlanResultViewerColumnFilterSpec> loadPlanResultViewerColumnFiltersForScope(
            String scopeKey) {
        try {
            if (scopeKey == null || scopeKey.isEmpty() || !Files.isRegularFile(STORE)) {
                return List.of();
            }
            JsonNode root = JSON.readTree(STORE.toFile());
            if (root == null || !root.isObject()) {
                return List.of();
            }
            JsonNode arr = root.get(scopeKey + SCOPE_SUFFIX_COLUMN_FILTERS);
            if (arr == null || !arr.isArray()) {
                return List.of();
            }
            List<PlanResultViewerColumnFilterSpec> out = new ArrayList<>();
            for (JsonNode n : arr) {
                if (n == null || !n.isObject()) {
                    continue;
                }
                String title = n.path("title").asText("");
                JsonNode vals = n.get("values");
                if (vals == null || !vals.isArray()) {
                    continue;
                }
                List<String> vs = new ArrayList<>();
                for (JsonNode v : vals) {
                    vs.add(v != null && v.isTextual() ? v.asText("") : v.asText(""));
                }
                if (!title.isEmpty()) {
                    out.add(new PlanResultViewerColumnFilterSpec(title, vs));
                }
            }
            return out;
        } catch (IOException e) {
            return List.of();
        }
    }

    /** 計画結果ビューアの列フィルタ状態を {@link #planResultViewerSheetScopeKey} に保存する。 */
    public static void savePlanResultViewerColumnFiltersForScope(
            String scopeKey, SpreadsheetView view, List<String> headersInOrder) {
        if (scopeKey == null || scopeKey.isEmpty() || view == null) {
            return;
        }
        try {
            Map<Integer, Set<String>> byIndex =
                    SpreadsheetMultiColumnFilterCoordinator.copyColumnAllowedByIndex(view);
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
            for (Map.Entry<Integer, Set<String>> e : byIndex.entrySet()) {
                int idx = e.getKey();
                if (headersInOrder == null
                        || idx < 0
                        || idx >= headersInOrder.size()) {
                    continue;
                }
                String title = headersInOrder.get(idx);
                ObjectNode o = JSON.createObjectNode();
                o.put("title", title != null ? title : "");
                ArrayNode vals = JSON.createArrayNode();
                for (String s : e.getValue()) {
                    vals.add(s != null ? s : "");
                }
                o.set("values", vals);
                arr.add(o);
            }
            root.set(scopeKey + SCOPE_SUFFIX_COLUMN_FILTERS, arr);
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
                                    while (c.next()) {
                                        /*
                                         * setGrid などで suppress 中に列が差し替わると、ここで全体 return すると
                                         * 新しい SpreadsheetColumn に幅リスナが付かず、以後の列幅が永続化されない。
                                         * 追加列へのリスナ装着は suppress に関わらず行い、保存だけ抑制する。
                                         */
                                        if (c.wasAdded()) {
                                            for (SpreadsheetColumn col : c.getAddedSubList()) {
                                                attachSpreadsheetWidthDebounced(
                                                        col, scheduleWidthSave, suppressSave);
                                            }
                                        }
                                        if (!suppressSave.getAsBoolean()
                                                && (c.wasPermutated()
                                                        || c.wasAdded()
                                                        || c.wasRemoved())) {
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

    /**
     * ポータブル自動バージョンアップ直後に、{@code pm-ai-data/config/bundled_table_column_order.json}（またはクラスパス
     * 既定）の内容で {@link #STORE} を上書きする。既存の列順・幅設定は失われる。
     */
    public static void overwriteTableColumnOrderStoreFromBundledAfterPortableUpgrade() throws IOException {
        JsonNode bundled = readBundledTableColumnOrderRoot();
        if (bundled == null || !bundled.isObject()) {
            return;
        }
        Files.createDirectories(STORE.getParent());
        JSON.writerWithDefaultPrettyPrinter().writeValue(STORE.toFile(), ((ObjectNode) bundled).deepCopy());
    }

    /**
     * ポータルアップデート直後: リポジトリ {@code init_setting/table_column_defaults.json}（配布同期で上書き済み）があれば
     * それを {@link #STORE} に書き、無ければ {@link #portableBundleInitSettingDir()}、最後にバンドルマージへフォールバックする。
     */
    public static void overwriteTableColumnOrderStoreAfterPortableUpgrade(Map<String, String> ui) throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path repo =
                InitSettingPaths.resolveRepoInitSettingDir(u).resolve(InitSettingPaths.TABLE_COLUMN_DEFAULTS_FILE);
        if (tryWriteTableColumnStoreFromPath(repo)) {
            return;
        }
        Path portable =
                InitSettingPaths.portableBundleInitSettingDir().resolve(InitSettingPaths.TABLE_COLUMN_DEFAULTS_FILE);
        if (tryWriteTableColumnStoreFromPath(portable)) {
            return;
        }
        overwriteTableColumnOrderStoreFromBundledAfterPortableUpgrade();
    }

    private static boolean tryWriteTableColumnStoreFromPath(Path file) throws IOException {
        if (!Files.isRegularFile(file)) {
            return false;
        }
        JsonNode n = JSON.readTree(file.toFile());
        if (n == null || !n.isObject()) {
            return false;
        }
        Files.createDirectories(STORE.getParent());
        JSON.writerWithDefaultPrettyPrinter().writeValue(STORE.toFile(), ((ObjectNode) n).deepCopy());
        return true;
    }

    /**
     * {@code table-column-order.json} が無いときだけ、バンドル既定をユーザーホームへ書き出す。空の JSON オブジェクト
     * でもファイルを作り、以降の保存がマージで動作するようにする。
     */
    public static void materializeBundledDefaultsIfStoreMissing() {
        try {
            if (Files.isRegularFile(STORE)) {
                return;
            }
            JsonNode bundled = readBundledTableColumnOrderRoot();
            ObjectNode root =
                    bundled != null && bundled.isObject()
                            ? (ObjectNode) bundled.deepCopy()
                            : JSON.createObjectNode();
            Files.createDirectories(STORE.getParent());
            JSON.writerWithDefaultPrettyPrinter().writeValue(STORE.toFile(), root);
        } catch (IOException ignored) {
        }
    }

    /**
     * 工場 UI リセット直後（ユーザーホームの {@link #STORE} を削除したあと）に呼ぶ。リポジトリの {@code
     * init_setting/table_column_defaults.json} があればグローバル設定で書き出した内容を {@link #STORE} へ複製し、無ければ
     * {@link #materializeBundledDefaultsIfStoreMissing()} にフォールバックする。
     */
    public static void materializeTableColumnStoreAfterFactoryReset(Map<String, String> ui) {
        try {
            Path exported =
                    InitSettingPaths.resolveRepoInitSettingDir(ui != null ? ui : Map.of())
                            .resolve(InitSettingPaths.TABLE_COLUMN_DEFAULTS_FILE);
            if (Files.isRegularFile(exported)) {
                JsonNode n = JSON.readTree(exported.toFile());
                if (n != null && n.isObject()) {
                    Files.createDirectories(STORE.getParent());
                    JSON.writerWithDefaultPrettyPrinter().writeValue(STORE.toFile(), n);
                    return;
                }
            }
        } catch (IOException ignored) {
        }
        materializeBundledDefaultsIfStoreMissing();
    }

    /**
     * 優先順（後勝ち）: クラスパス {@code bundled_table_column_order.json} → {@code pm-ai-data/config} 同名 →
     * {@code init_setting/table_column_defaults.json} → {@code pm-ai-data/init_setting/table_column_defaults.json}。
     */
    private static JsonNode readBundledTableColumnOrderRoot() {
        ObjectNode acc = JSON.createObjectNode();
        mergeTableColumnRootFromClasspath(acc);
        try {
            Path beside =
                    Path.of(System.getProperty("user.dir", "."))
                            .toAbsolutePath()
                            .normalize()
                            .resolve("pm-ai-data")
                            .resolve("config")
                            .resolve("bundled_table_column_order.json");
            mergeTableColumnRootFromPath(acc, beside);
        } catch (Exception ignored) {
        }
        mergeTableColumnRootFromPath(
                acc,
                InitSettingPaths.cwdInitSettingDir().resolve(InitSettingPaths.TABLE_COLUMN_DEFAULTS_FILE));
        mergeTableColumnRootFromPath(
                acc,
                InitSettingPaths.portableBundleInitSettingDir()
                        .resolve(InitSettingPaths.TABLE_COLUMN_DEFAULTS_FILE));
        return acc.size() > 0 ? acc : null;
    }

    private static void mergeTableColumnRootFromClasspath(ObjectNode acc) {
        try (InputStream in =
                TableColumnOrderPersistence.class.getResourceAsStream(BUNDLED_TABLE_COLUMN_ORDER_RESOURCE)) {
            if (in == null) {
                return;
            }
            JsonNode n = JSON.readTree(in);
            if (n != null && n.isObject()) {
                deepMergeTableColumnRoot(acc, (ObjectNode) n);
            }
        } catch (IOException ignored) {
        }
    }

    private static void mergeTableColumnRootFromPath(ObjectNode acc, Path file) {
        try {
            if (!Files.isRegularFile(file)) {
                return;
            }
            JsonNode n = JSON.readTree(file.toFile());
            if (n != null && n.isObject()) {
                deepMergeTableColumnRoot(acc, (ObjectNode) n);
            }
        } catch (IOException ignored) {
        }
    }

    private static void deepMergeTableColumnRoot(ObjectNode acc, ObjectNode overlay) {
        Iterator<String> fn = overlay.fieldNames();
        while (fn.hasNext()) {
            String k = fn.next();
            JsonNode v = overlay.get(k);
            if (v != null) {
                acc.set(k, v.deepCopy());
            }
        }
    }
}
