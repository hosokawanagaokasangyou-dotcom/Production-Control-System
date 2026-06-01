package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Comparator;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.HashSet;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

import javafx.beans.property.ObjectProperty;
import javafx.beans.property.ReadOnlyObjectWrapper;
import javafx.beans.property.SimpleObjectProperty;
import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Node;
import javafx.scene.Scene;
import javafx.scene.text.Font;
import javafx.scene.text.Text;
import javafx.scene.text.TextFlow;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.ListCell;
import javafx.scene.control.ListView;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.TitledPane;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.control.Tab;
import javafx.scene.control.TabPane;
import javafx.scene.control.TextField;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableView;
import javafx.scene.control.TablePosition;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.TextFormatter;
import javafx.scene.control.TextInputDialog;
import javafx.scene.control.Tooltip;
import javafx.scene.input.Clipboard;
import javafx.scene.input.ClipboardContent;
import javafx.scene.input.DragEvent;
import javafx.scene.input.Dragboard;
import javafx.scene.input.MouseEvent;
import javafx.scene.input.TransferMode;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.stage.Modality;
import javafx.stage.Screen;
import javafx.stage.Stage;
import javafx.stage.Window;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.GridChange;
import org.controlsfx.control.spreadsheet.SpreadsheetCell;
import org.controlsfx.control.spreadsheet.SpreadsheetCellType;
import org.controlsfx.control.spreadsheet.SpreadsheetColumn;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup;
import jp.co.pm.ai.desktop.dispatch.AladdinSystemDispatchDisplayQty;
import jp.co.pm.ai.desktop.dispatch.DispatchAladdinPlanAligner;
import jp.co.pm.ai.desktop.dispatch.rules.ui.trace.DispatchRuleApplicationBadgeSupport;
import jp.co.pm.ai.desktop.dispatch.DispatchInteractiveDateAxis;
import jp.co.pm.ai.desktop.dispatch.DispatchInteractiveRollUnitSupport;
import jp.co.pm.ai.desktop.dispatch.DispatchTrialConsistency;
import jp.co.pm.ai.desktop.dispatch.DispatchTrialShortages;
import jp.co.pm.ai.desktop.dispatch.DispatchTrialShortages.DispatchQtyShortfallRow;
import jp.co.pm.ai.desktop.dispatch.DispatchTimelineCalendarMetersIndex;
import jp.co.pm.ai.desktop.dispatch.DispatchTimelineMetaMissShortfalls;
import jp.co.pm.ai.desktop.dispatch.DispatchPlanInputInteractiveCoverageCheck;
import jp.co.pm.ai.desktop.dispatch.DispatchPlanInputInteractiveCoverageCheck.TaskKey;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchDocument;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchInteractiveGridModel;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchJsonIo;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchNormalizer;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchPivot;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchPythonExport;
import jp.co.pm.ai.desktop.dispatch.RawInputMorningDispatchRateWarning;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchSchema;
import jp.co.pm.ai.desktop.dispatch.OvertimeSimulationOverridesReader;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchStage3Support;
import jp.co.pm.ai.desktop.dispatch.Stage3PlanningMetaStore;
import jp.co.pm.ai.desktop.dispatch.Stage21AttendanceApplyStatusPanel;
import jp.co.pm.ai.desktop.dispatch.Stage21TrialSnapshotStore;
import jp.co.pm.ai.desktop.dispatch.Stage3DispatchQtyBalanceCheck;
import jp.co.pm.ai.desktop.ui.TabularCellHighlight;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchStage2ColumnSupport;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchTrialPython;
import jp.co.pm.ai.desktop.dispatch.ResultTaskUnassignedLoader;
import jp.co.pm.ai.desktop.io.JsonTableIo;
import jp.co.pm.ai.desktop.ui.ColumnVisibilitySupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnDragReorderSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnReorderDialog;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnSettingsStrip;
import jp.co.pm.ai.desktop.ui.SpreadsheetMultiColumnFilterCoordinator;
import jp.co.pm.ai.desktop.ui.SpreadsheetRowReorderDragGhost;
import jp.co.pm.ai.desktop.ui.SpreadsheetTabularSupport;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence.DispatchInteractiveDateAxisPastDaysPrefs;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence.DispatchInteractiveDateQtyLineFilterPrefs;
import jp.co.pm.ai.planning.stage2.core.Stage2PlanRowDispatchQtyMetrics;
import jp.co.pm.ai.planning.stage2.input.Stage2MasterFactoryHoursReader;
import jp.co.pm.ai.planning.stage2.core.Stage2RollUnitLengthTables;

/**
 * Interactive pivot editor for result dispatch JSON (ControlsFX SpreadsheetView: task-by-day columns +
 * process+machine-by-day).
 */
public final class DispatchInteractiveTabController {

    /** 起動直後は「実行・ログ」タブが選択されるため、初回 JSON 読込は配台タブ初回表示まで遅延する。 */
    private final AtomicBoolean pendingInitialDispatchReload = new AtomicBoolean(false);

    /** オフシーンで JSON だけ読み込んだあと、配台タブ表示時にグリッド再構築する。 */
    private final AtomicBoolean pendingGridRebuildAfterTabAttach = new AtomicBoolean(false);


    private record ReloadBundle(
            ResultDispatchDocument doc, DispatchTimelineCalendarMetersIndex timelineMeters) {}

    private record DispatchSaveOutcome(Path jsonPath, String xlsxStdoutLine) {}

    private static final String DND_PREFIX = "pm-dispatch-dnd|wide|";
    private static final String DND_V2_MARKER = "v2|";
    /** Drag payload for reordering wide-grid profile rows (leading columns only). */
    private static final String DND_ROW_PREFIX = "pm-dispatch-dnd|wide|row|v2|";

    /** One spreadsheet column per calendar day (wide grid date axis). */
    private static final int DAY_SLOT_COLUMNS = 1;

    /**
     * Width for date columns where every data cell is calendar-blocked (gray); keeps holiday bands from wasting
     * horizontal space ({@code SpreadsheetView} uses pixels; value is converted from typographic points).
     */
    private static final double BLOCKED_DATE_COLUMN_PREF_PT = 5.0;

    /**
     * Minimum width for normal date columns so {@code YYYY-MM-DD} plus ControlsFX filter header stays readable
     * (not truncated to "2026-...").
     */
    private static final double MIN_DATE_COLUMN_WIDTH_PX = 112.0;

    /** Fully-blocked (holiday) date columns: stay narrow but wide enough for a short header glyph. */
    private static final double MIN_BLOCKED_DATE_COLUMN_WIDTH_PX = 40.0;

    /** 日付列で数量が正のとき（薄緑・黒字。行選択時も読めるよう text-fill を明示）。 */
    private static final String DATE_CELL_STYLE_POSITIVE_QTY =
            "-fx-background-color: #d4edd4; -fx-text-fill: #111111;";

    /** 配台試行でタイムライン実績が目標に届かないセル。 */
    private static final String DATE_CELL_STYLE_SHORTFALL =
            "-fx-background-color: #b71c1c; -fx-text-fill: #ffffff; -fx-font-weight: bold;";

    /** アラジン整列: 数量が減った（移動元）日付セル。 */
    private static final String DATE_CELL_STYLE_ALADDIN_ALIGN_SOURCE =
            "-fx-background-color: #fef3c7; -fx-text-fill: #92400e;";

    /** アラジン整列: 数量が増えた（移動先）日付セル。 */
    private static final String DATE_CELL_STYLE_ALADDIN_ALIGN_DEST =
            "-fx-background-color: #bfdbfe; -fx-text-fill: #111111;";

    private static final String DISPATCH_DATE_QTY_CELL_STYLE_CLASS = "dispatch-date-qty-cell";

    private static final String DISPATCH_DATE_QTY_SHORTFALL_CELL_STYLE_CLASS = "dispatch-date-qty-shortfall-cell";

    /** 段階2.1 試行で数量が変わった日付セル。 */
    private static final String DISPATCH_STAGE21_QTY_CHANGED_CELL_STYLE_CLASS =
            "dispatch-stage21-qty-changed-cell";

  /** 段階3の2行表示（{@code \\n}）用。CSS の wrap-text は付けない。 */
    private static final String DISPATCH_DATE_QTY_MULTILINE_CELL_STYLE_CLASS =
            "dispatch-date-qty-multiline";

    /** 日付セル内の (段階3後) 行（{@link #LABEL_STAGE3_ACTUAL} で始まる行）用。 */
    private static final String DISPATCH_STAGE3_AFTER_LINE_STYLE_CLASS = "dispatch-stage3-after-line";

    /** 段階3試行後に当日配台数量を手動変更したセルの (段階3改) 行用。 */
    private static final String DISPATCH_STAGE3_REVISED_LINE_STYLE_CLASS = "dispatch-stage3-revised-line";

    /**
     * 段階3ラベル付き Graphic 表示時。内側 TextField の素文字列が手前に重なり CSS が効かないのを防ぐ。
     */
    private static final String DISPATCH_DATE_QTY_GRAPHIC_ONLY_STYLE_CLASS =
            "dispatch-date-qty-graphic-only";

    /** ControlsFX の loadStyle 上書き対策で Label にも直接指定（(段階3改) 青・太字）。 */
    private static final String STAGE3_REVISED_LINE_INLINE_STYLE =
            "-fx-font-weight: bold; -fx-text-fill: #1565C0;";

    private static final String STAGE3_AFTER_LINE_INLINE_STYLE =
            "-fx-font-weight: bold; -fx-text-fill: #111111;";

    /** (アラ計画)/(段階3前) 等の通常行。行選択時の白字化を防ぐ。 */
    private static final String STAGE3_QTY_DEFAULT_LINE_INLINE_STYLE = "-fx-text-fill: #111111;";

    private static final String STAGE3_QTY_DEFAULT_TEXT_INLINE_STYLE = "-fx-fill: #111111;";

    /** 日付セル表示: 段階3試行後の手動改定（当日配台数量）。 */
    static final String LABEL_STAGE3_REVISED = "(\u6bb5\u968e3\u6539)";

    /** 2行ラベル付き数量セルの行高（px）。 */
    private static final double DISPATCH_STAGE3_MULTILINE_ROW_HEIGHT_PX = 44.0;

    /** (アラ計画)＋(段階3前)＋(段階3後/改) の3行固定表示用行高（px）。 */
    private static final double DISPATCH_ALADDIN_STAGE3_MULTILINE_ROW_HEIGHT_PX = 66.0;

    /** 段階2.1 比較表示: 上記3行＋(段階2.1後) の4行固定表示用行高（px）。 */
    private static final double DISPATCH_ALADDIN_STAGE21_MULTILINE_ROW_HEIGHT_PX = 88.0;

    /** 段階2.1試行後（段階3未実行）: (アラ計画)+(段階2後)+(段階2.1後) の3行。 */
    private static final double DISPATCH_ALADDIN_STAGE21_PRE_STAGE3_MULTILINE_ROW_HEIGHT_PX = 66.0;

    /** 段階3日付セル: 1行目=(アラ計画)、2行目=(段階3前)、3行目=(段階3後)または(段階3改)。 */
    private static final int STAGE3_QTY_FIXED_LINE_COUNT = 3;

    /** 段階2.1試行後（段階3実行済）: 4行目=(段階2.1後)。 */
    private static final int STAGE21_QTY_FIXED_LINE_COUNT = 4;

    private static final int STAGE21_PRE_STAGE3_QTY_LINE_COUNT = 3;

    /**
     * 段階3数量セルの固定行スロット。{@link #visible} が false の行は {@code \u00a0} で行位置だけ確保する。
     */
    record Stage3QtyLineSlot(String lineText, boolean visible) {}

    /** 「アラジン計画に合わせる」直後のセル移動ハイライト（再読み・次回整列で消える）。 */
    private enum AladdinAlignMoveHighlight {
        SOURCE,
        DEST
    }

    private static final DateTimeFormatter ALADDIN_PLAN_DATE_FMT =
            DateTimeFormatter.ofPattern("yyyy/MM/dd");

    /** 日付セル表示: アラジン加工計画の計画数量（納期管理ビューと同じ {@code (アラ計画)} 表記）。 */
    static final String LABEL_ALADDIN_PLAN = "(\u30a2\u30e9\u8a08\u753b)";

    /** 日付セル表示: 段階2成果の配台数量（段階3試行前・当日配台数量）。 */
    static final String LABEL_STAGE2_PLAN = "(\u6bb5\u968e2\u5f8c)";

    /** 日付セル表示: 段階3試行前の編集対象（当日配台数量）。 */
    static final String LABEL_STAGE3_PLAN = "(\u6bb5\u968e3\u524d)";

    /** 日付セル表示: 段階3試行後の実績（実配台数量・DnD/編集では変更しない）。 */
    static final String LABEL_STAGE3_ACTUAL = "(\u6bb5\u968e3\u5f8c)";

    /** 日付セル表示: 段階2.1試行後の配台（stage21 タイムライン暦日m）。 */
    static final String LABEL_STAGE21_ACTUAL = "(\u6bb5\u968e2.1\u5f8c)";

    /** 日付セル内の (段階2.1後) 行用。 */
    private static final String DISPATCH_STAGE21_AFTER_LINE_STYLE_CLASS = "dispatch-stage21-after-line";

    private static final String STAGE21_AFTER_LINE_INLINE_STYLE =
            "-fx-font-weight: bold; -fx-text-fill: #2E7D32;";

    /** 段階3試行後の実配台数量（日付セルの括弧内）を行単位で合算する列。 */
    private static final String COL_STAGE3_DISPATCH_QTY_TOTAL = "段階3配台数";

    private static final String COL_SPECIAL_RULES = "特別ルール";

    private static final List<String> WIDE_STATIC_HEADERS =
            List.of(
                    ResultDispatchSchema.COL_DISPATCH_TRIAL_ORDER,
                    ResultDispatchSchema.COL_PROCESS,
                    ResultDispatchSchema.COL_MACHINE,
                    "加工内容",
                    ResultDispatchSchema.COL_ORDER_NO,
                    "依頼NO",
                    "換算数量",
                    "実加工数",
                    "計画合計",
                    COL_STAGE3_DISPATCH_QTY_TOTAL,
                    Stage3DispatchQtyBalanceCheck.COL_TITLE,
                    COL_SPECIAL_RULES);

    /** 「工程+機械×日」ビューの先頭固定列（日付ブロックの直前まで）。 */
    private static final List<String> BY_DAY_STATIC_HEADERS =
            List.of(
                    ResultDispatchSchema.COL_PROCESS,
                    ResultDispatchSchema.COL_MACHINE,
                    "加工内容",
                    COL_STAGE3_DISPATCH_QTY_TOTAL);

    private record WideGridBundle(
            GridBase grid,
            List<Map<String, String>> profiles,
            List<WideRow> rowItems,
            boolean[] blockedCols,
            int staticCols,
            int dayCount) {}

    private record ByDayGridBundle(GridBase grid, boolean[] blockedCols, int staticCols, int dayCount) {}

    private record FullGridRebuild(List<LocalDate> axis, WideGridBundle wide, ByDayGridBundle byDay) {}

    /**
     * In-memory {@link #doc} differs from last successful「保存」to disk（行順・セル編集・DnD 等）。
     * 配台試行は開始時に表を JSON へ自動同期するため、試行ボタンは未保存でも押下可能。
     */
    private boolean dispatchDocDirtySinceSave;

    /** True while load/rebuild progress UI disables the toolbar ({@link #setReloadInteractionDisabled}). */
    private boolean reloadInteractionDisabled;

    /** 納期管理ビュー再読み込み中（メインシェルから同期）。 */
    private boolean deliveryCalendarReloadBlocking;

    /** 段階1／段階2／段階2.1 パイプライン実行中。 */
    private boolean stagePipelineBusy;

    /** 入力3表生成 Python 実行中（ボタン無効・プログレス表示）。 */
    private boolean stage3InputBuildBusy;

    /** Avoid treating programmatic grid updates as user edits ({@link #onWideGridChange}). */
    private final AtomicBoolean suppressDispatchGridDirty = new AtomicBoolean(false);

    /**
     * {@link #rebuildGrids} や空表示クリアのたびに増やし、遅延した列同期・可視化ジョブが古いグリッドに触れないようにする。
     */
    private final AtomicInteger dispatchSpreadsheetLayoutGeneration = new AtomicInteger(0);

    /** shaped アラジン JSON 由来の計画数量ルックアップ（{@link #buildFullGridRebuild} で更新）。 */
    private Map<String, Map<String, Map<String, Map<String, Double>>>> aladdinPlanLookup = Map.of();

    @FXML
    private ProgressIndicator busyIndicator;

    @FXML
    private Button loadButton;

    @FXML
    private Button saveButton;

    @FXML
    private Button dispatchTrialButton;

    @FXML
    private Button buildStage3InputButton;

    @FXML
    private static final String DISPATCH_TRIAL_BUTTON_TEXT_DEFAULT = "段階3";

    private static final String DISPATCH_TRIAL_BUTTON_TEXT_SUMMARY_LOCKED =
            "段階3（サマリエクセル更新中）";

    private static final String DISPATCH_TRIAL_BUTTON_TEXT_DELIVERY_CALENDAR_RELOAD =
            "段階3（納期管理ビュー更新中）";

    private static final String ALIGN_TO_ALADDIN_PLAN_BUTTON_TEXT_DEFAULT = "アラジン計画に合わせる";

    private static final String ALIGN_TO_ALADDIN_PLAN_BUTTON_TEXT_DELIVERY_CALENDAR_RELOAD =
            "納期管理ビュー更新中";

    private static final String BUILD_STAGE3_INPUT_BUTTON_TEXT_DEFAULT = "入力3表を生成";

    private static final String BUILD_STAGE3_INPUT_BUTTON_TEXT_BUSY = "生成中…";

    private static final String BUILD_STAGE3_INPUT_STATUS_BUSY = "入力3表を生成中…";

    @FXML
    private Button wideRowUpButton;

    @FXML
    private Button wideRowDownButton;

    @FXML
    private Button dispatchReplayButton;

    @FXML
    private Button dispatchReplayStopButton;

    @FXML
    private javafx.scene.control.Slider dispatchReplaySpeedSlider;

    /** 配台リプレイ（プラン B）。グリッドモデルは触らず、ワイド表のセル選択/スクロールのみで再生する。 */
    private final jp.co.pm.ai.desktop.dispatch.DispatchReplayController dispatchReplay =
            new jp.co.pm.ai.desktop.dispatch.DispatchReplayController();

    @FXML
    private CheckBox showAladdinPlanQtyLineCheck;

    @FXML
    private CheckBox showStage2PlanQtyLineCheck;

    @FXML
    private CheckBox showStage3PlanQtyLineCheck;

    @FXML
    private CheckBox showStage3AfterQtyLineCheck;

    @FXML
    private CheckBox showStage21AfterQtyLineCheck;

    @FXML
    private Spinner<Integer> dateAxisPastDaysSpinner;

    @FXML
    private Button alignToAladdinPlanButton;

    @FXML
    private Label statusLabel;

    @FXML
    private ProgressBar reloadProgressBar;

    @FXML
    private Label jsonPathLabel;

    @FXML
    private Label dispatchPlanningStageBadgeLabel;

    @FXML
    private VBox stage21AttendanceApplyPanel;

    @FXML
    private Label stage21AttendanceHeadlineLabel;

    @FXML
    private Label stage21AttendanceSummaryLabel;

    @FXML
    private Label stage21AttendanceOverridesLabel;

    @FXML
    private Label stage21AttendancePythonLabel;

    @FXML
    private Label stage21AttendanceAppliedAtLabel;

    @FXML
    private VBox dispatchShortfallPanel;

    @FXML
    private TableView<DispatchQtyShortfallRow> dispatchShortfallTable;

    @FXML
    private TabPane innerTabPane;

    private final AtomicBoolean suppressInnerTabSessionPersistence = new AtomicBoolean(false);

    private volatile boolean innerTabPersistenceWired;

    @FXML
    private StackPane wideSpreadsheetHost;

    @FXML
    private HBox wideColumnStripHost;

    @FXML
    private StackPane byDaySpreadsheetHost;

    @FXML
    private HBox byDayColumnStripHost;

    private final SpreadsheetView wideSpreadsheet = new SpreadsheetView();
    private final SpreadsheetView byDaySpreadsheet = new SpreadsheetView();

    private MainShellController shell;

    private ResultDispatchDocument doc = ResultDispatchDocument.empty();

    /** 設備ガント契約由来の暦日別加工量(m)。タスク×日付の段階3表示をガントと揃える。 */
    private DispatchTimelineCalendarMetersIndex timelineCalendarMeters =
            DispatchTimelineCalendarMetersIndex.empty();

    /** 段階2.1 成果物（output/stage21/）のタイムライン暦日m。(段階2.1後) 表示用。 */
    private DispatchTimelineCalendarMetersIndex stage21TimelineCalendarMeters =
            DispatchTimelineCalendarMetersIndex.empty();

    private volatile boolean retainStage21TrialMetaOnNextReload;

    private List<LocalDate> dateAxis = new ArrayList<>();

    /**
     * 配台試行直後の {@link ResultDispatchSchema#COL_DISPATCH_QTY} 合計（プロファイル×配台日）。
     * 手動移動・編集でずれたセルは (段階3改) 表示に切り替える。
     */
    private final Map<String, Double> stage3TrialPlanQtySnapshot = new HashMap<>();

    /**
     * 段階3試行直前の結果_配台表（段階2後＋手動修正）。段階3.5 の Python 入力の正とする。
     */
    private ResultDispatchDocument stage3TrialInputDocumentSnapshot;

    /**
     * 段階2.1試行直前の段階2配台数量（プロファイル×配台日）。sidecar がある間は
     * (段階2後) と (段階2.1後) を比較表示する。
     */
    private final Map<String, Double> stage21BaselinePlanQtySnapshot = new HashMap<>();

    /**
     * 「アラジン計画に合わせる」で数量が変わったセル（キーは {@link DispatchTrialShortages#wideShortfallKey}）。
     */
    private final Map<String, AladdinAlignMoveHighlight> aladdinAlignMoveHighlights =
            new HashMap<>();

    private boolean stage21TrialApplied;

    private Stage21TrialSnapshotStore.Stage21TrialMeta stage21TrialMeta =
            Stage21TrialSnapshotStore.Stage21TrialMeta.empty();

    private String lastShortagesNote = "";

    private boolean pendingStage3TrialSnapshotCapture;

    /** 段階3.0/3.1/3.2 直後の再読込で baseline snapshot を上書きしない。 */
    private boolean retainStage3TrialSnapshotOnNextReload;

    /**
     * 日付列のユーザー希望順（両グリッド共通）。{@code null} のときは {@link #computeDateAxisList()} の自然順を使う。
     */
    private List<LocalDate> preferredDateAxisOrder;

    /** 列ドラッグ並べ替え起因の {@link #rebuildGrids()} 中はヘッダ変更コールバックを無視する。 */
    private final AtomicBoolean suppressColumnReorderPersistence = new AtomicBoolean(false);

    /** グリッド再構築時の列幅復元中は {@link TableColumnOrderPersistence} への保存を抑止する。 */
    private final AtomicBoolean suppressDispatchColumnLayoutPersistence = new AtomicBoolean(false);

    private boolean dispatchColumnLayoutWatchersInstalled;

    /** 横スクロール時に固定する先頭列数（タスク×日付）。 */
    private final AtomicInteger headerColumnCountWide = new AtomicInteger(0);

    /** 横スクロール時に固定する先頭列数（工程+機械×日）。 */
    private final AtomicInteger headerColumnCountByDay = new AtomicInteger(0);

    private final List<Map<String, String>> wideProfiles = new ArrayList<>();

    /** Parallel to {@link #wideProfiles} rows in the wide grid. */
    private final List<WideRow> wideRowItems = new ArrayList<>();

    private final DispatchRuleApplicationBadgeSupport specialRuleBadges =
            new DispatchRuleApplicationBadgeSupport();

    private List<DispatchQtyShortfallRow> lastDispatchShortfallRows = List.of();

    /** {@link DispatchTrialShortages.FullBundle#shortageHints()}（op_shortage / as_shortage）。試行後ダイアログ用。 */
    private List<DispatchTrialShortages.ShortageHint> lastDispatchShortageHints = List.of();

    private final Set<String> dispatchWideShortfallKeys = new HashSet<>();

    private final Set<String> dispatchByDayShortfallKeys = new HashSet<>();

    private final AtomicReference<Stage2RollUnitLengthTables> cachedRollUnitTables =
            new AtomicReference<>();

    private DispatchInteractiveDateQtyLineFilterPrefs dateQtyLineFilter =
            DispatchInteractiveDateQtyLineFilterPrefs.defaults();

    /** 結果_配台表 sidecar または旧 JSON から解決した段階3.0/3.1/3.2（未実行時 NONE）。 */
    private ResultDispatchStage3Support.Stage3PlanningVariant stage3PlanningVariant =
            ResultDispatchStage3Support.Stage3PlanningVariant.NONE;

    private int dateAxisPastDays =
            TableColumnOrderPersistence.loadDispatchInteractiveDateAxisPastDaysPrefs().pastDays();

    private final AtomicBoolean suppressDateQtyLineFilterUi = new AtomicBoolean(false);

    private final AtomicBoolean suppressDateAxisPastDaysUi = new AtomicBoolean(false);

    @FXML
    private void initialize() {
        StackPane.setAlignment(wideSpreadsheet, Pos.TOP_LEFT);
        wideSpreadsheetHost.getChildren().setAll(wideSpreadsheet);
        VBox.setVgrow(wideSpreadsheetHost, javafx.scene.layout.Priority.ALWAYS);
        wideSpreadsheet.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
        wideSpreadsheet.prefHeightProperty().bind(wideSpreadsheetHost.heightProperty());

        StackPane.setAlignment(byDaySpreadsheet, Pos.TOP_LEFT);
        byDaySpreadsheetHost.getChildren().setAll(byDaySpreadsheet);
        VBox.setVgrow(byDaySpreadsheetHost, javafx.scene.layout.Priority.ALWAYS);
        byDaySpreadsheet.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
        byDaySpreadsheet.prefHeightProperty().bind(byDaySpreadsheetHost.heightProperty());

        // 固定配色（薄緑日付列・黒字）。ThemeBridge は -fx-text-inner-color で日付セルが白抜けするため入れない。
        SpreadsheetTabularSupport.installPmAiReadableSpreadsheetChrome(wideSpreadsheet);
        SpreadsheetTabularSupport.installPmAiReadableSpreadsheetChrome(byDaySpreadsheet);

        wideSpreadsheet.getSelectionModel().setSelectionMode(javafx.scene.control.SelectionMode.MULTIPLE);
        byDaySpreadsheet.getSelectionModel().setSelectionMode(javafx.scene.control.SelectionMode.MULTIPLE);
        SpreadsheetTabularSupport.installFullRowDataSelection(wideSpreadsheet);
        SpreadsheetTabularSupport.installFullRowDataSelection(byDaySpreadsheet);

        SpreadsheetTabularSupport.installSpreadsheetChromeRelayoutDebouncerForHost(
                wideSpreadsheetHost, this::resolvedWideLeadingColumnCount);
        SpreadsheetTabularSupport.installSpreadsheetChromeRelayoutDebouncerForHost(
                byDaySpreadsheetHost, this::resolvedByDayLeadingColumnCount);

        installWideDnDHandlers();
        installWideDoubleClickHandler();
        installByDayDoubleClickHandler();
        ensureDispatchColumnStripsInstalled();
        ensureDispatchColumnLayoutWatchersInstalled();

        if (dispatchShortfallTable != null) {
            installDispatchShortfallColumns(dispatchShortfallTable);
            dispatchShortfallTable.setColumnResizePolicy(
                    TableView.CONSTRAINED_RESIZE_POLICY_FLEX_LAST_COLUMN);
            wireDispatchShortfallSelectionToWideGrid();
        }
        wireDateQtyLineFilterControls();
        wireDateAxisPastDaysSpinner();
    }

    private void wireDateQtyLineFilterControls() {
        dateQtyLineFilter = TableColumnOrderPersistence.loadDispatchInteractiveDateQtyLineFilterPrefs();
        suppressDateQtyLineFilterUi.set(true);
        try {
            if (showAladdinPlanQtyLineCheck != null) {
                showAladdinPlanQtyLineCheck.setSelected(dateQtyLineFilter.showAladdinPlan());
            }
            if (showStage2PlanQtyLineCheck != null) {
                showStage2PlanQtyLineCheck.setSelected(dateQtyLineFilter.showStage2Plan());
            }
            if (showStage3PlanQtyLineCheck != null) {
                showStage3PlanQtyLineCheck.setSelected(dateQtyLineFilter.showStage3Plan());
            }
            if (showStage3AfterQtyLineCheck != null) {
                showStage3AfterQtyLineCheck.setSelected(dateQtyLineFilter.showStage3After());
            }
            if (showStage21AfterQtyLineCheck != null) {
                showStage21AfterQtyLineCheck.setSelected(dateQtyLineFilter.showStage21After());
            }
        } finally {
            suppressDateQtyLineFilterUi.set(false);
        }
        Runnable onFilterChanged =
                () -> {
                    if (suppressDateQtyLineFilterUi.get()) {
                        return;
                    }
                    dateQtyLineFilter = snapshotDateQtyLineFilterPrefs();
                    TableColumnOrderPersistence.saveDispatchInteractiveDateQtyLineFilterPrefs(
                            dateQtyLineFilter);
                    if (doc != null && !dateAxis.isEmpty()) {
                        rebuildGrids();
                    }
                };
        if (showAladdinPlanQtyLineCheck != null) {
            showAladdinPlanQtyLineCheck.selectedProperty().addListener((o, a, b) -> onFilterChanged.run());
        }
        if (showStage2PlanQtyLineCheck != null) {
            showStage2PlanQtyLineCheck.selectedProperty().addListener((o, a, b) -> onFilterChanged.run());
        }
        if (showStage3PlanQtyLineCheck != null) {
            showStage3PlanQtyLineCheck.selectedProperty().addListener((o, a, b) -> onFilterChanged.run());
        }
        if (showStage3AfterQtyLineCheck != null) {
            showStage3AfterQtyLineCheck.selectedProperty().addListener((o, a, b) -> onFilterChanged.run());
        }
        if (showStage21AfterQtyLineCheck != null) {
            showStage21AfterQtyLineCheck.selectedProperty().addListener((o, a, b) -> onFilterChanged.run());
        }
    }

    private void wireDateAxisPastDaysSpinner() {
        if (dateAxisPastDaysSpinner == null) {
            return;
        }
        DispatchInteractiveDateAxisPastDaysPrefs loaded =
                TableColumnOrderPersistence.loadDispatchInteractiveDateAxisPastDaysPrefs();
        dateAxisPastDays = loaded.pastDays();
        suppressDateAxisPastDaysUi.set(true);
        try {
            configureDateAxisPastDaysSpinner(dateAxisPastDaysSpinner, dateAxisPastDays);
        } finally {
            suppressDateAxisPastDaysUi.set(false);
        }
        Runnable onChanged =
                () -> {
                    if (suppressDateAxisPastDaysUi.get()) {
                        return;
                    }
                    commitDateAxisPastDaysSpinnerValue();
                    int next = resolveDateAxisPastDaysFromSpinner();
                    if (next == dateAxisPastDays) {
                        return;
                    }
                    dateAxisPastDays = next;
                    TableColumnOrderPersistence.saveDispatchInteractiveDateAxisPastDaysPrefs(
                            new DispatchInteractiveDateAxisPastDaysPrefs(dateAxisPastDays));
                    preferredDateAxisOrder = null;
                    if (doc != null) {
                        rebuildGrids();
                    }
                };
        dateAxisPastDaysSpinner.valueProperty().addListener((o, a, b) -> onChanged.run());
        dateAxisPastDaysSpinner
                .getEditor()
                .focusedProperty()
                .addListener(
                        (obs, was, is) -> {
                            if (Boolean.FALSE.equals(is)) {
                                onChanged.run();
                            }
                        });
    }

    private void commitDateAxisPastDaysSpinnerValue() {
        if (dateAxisPastDaysSpinner == null || !dateAxisPastDaysSpinner.isEditable()) {
            return;
        }
        try {
            dateAxisPastDaysSpinner.commitValue();
        } catch (IllegalArgumentException ex) {
            Integer cur = dateAxisPastDaysSpinner.getValue();
            if (cur != null) {
                dateAxisPastDaysSpinner.getEditor().setText(Integer.toString(cur));
            }
        }
    }

    private int resolveDateAxisPastDaysFromSpinner() {
        if (dateAxisPastDaysSpinner == null) {
            return DispatchInteractiveDateAxisPastDaysPrefs.DEFAULT_PAST;
        }
        String raw = dateAxisPastDaysSpinner.getEditor().getText();
        if (raw != null) {
            String t = raw.strip();
            if (!t.isEmpty()) {
                String digitsOnly = t.replaceAll("[^0-9]", "");
                if (!digitsOnly.isEmpty()) {
                    try {
                        int v = Integer.parseInt(digitsOnly);
                        return Math.max(
                                DispatchInteractiveDateAxisPastDaysPrefs.MIN,
                                Math.min(DispatchInteractiveDateAxisPastDaysPrefs.MAX, v));
                    } catch (NumberFormatException ignored) {
                        // fall through
                    }
                }
            }
        }
        Integer val = dateAxisPastDaysSpinner.getValue();
        return val != null ? val : DispatchInteractiveDateAxisPastDaysPrefs.DEFAULT_PAST;
    }

    private static void configureDateAxisPastDaysSpinner(Spinner<Integer> spinner, int initial) {
        int min = DispatchInteractiveDateAxisPastDaysPrefs.MIN;
        int max = DispatchInteractiveDateAxisPastDaysPrefs.MAX;
        SpinnerValueFactory.IntegerSpinnerValueFactory vf =
                new SpinnerValueFactory.IntegerSpinnerValueFactory(
                        min, max, Math.max(min, Math.min(max, initial)));
        vf.setAmountToStepBy(1);
        spinner.setValueFactory(vf);
        TextFormatter<Integer> formatter =
                new TextFormatter<>(
                        vf.getConverter(),
                        vf.getValue(),
                        change -> {
                            String text = change.getControlNewText();
                            if (text.isEmpty()) {
                                return change;
                            }
                            return text.matches("\\d{0,3}") ? change : null;
                        });
        spinner.getEditor().setTextFormatter(formatter);
    }

    private DispatchInteractiveDateQtyLineFilterPrefs snapshotDateQtyLineFilterPrefs() {
        return new DispatchInteractiveDateQtyLineFilterPrefs(
                showAladdinPlanQtyLineCheck == null || showAladdinPlanQtyLineCheck.isSelected(),
                showStage2PlanQtyLineCheck == null || showStage2PlanQtyLineCheck.isSelected(),
                showStage3PlanQtyLineCheck == null || showStage3PlanQtyLineCheck.isSelected(),
                showStage3AfterQtyLineCheck == null || showStage3AfterQtyLineCheck.isSelected(),
                showStage21AfterQtyLineCheck == null || showStage21AfterQtyLineCheck.isSelected());
    }

    private ResultDispatchStage3Support.Stage3PlanningVariant effectiveStage3PlanningVariant() {
        if (stage3PlanningVariant != null
                && stage3PlanningVariant
                        != ResultDispatchStage3Support.Stage3PlanningVariant.NONE) {
            return stage3PlanningVariant;
        }
        if (docHasActualDispatchQtyColumn()) {
            return ResultDispatchStage3Support.Stage3PlanningVariant.LEGACY;
        }
        return ResultDispatchStage3Support.Stage3PlanningVariant.NONE;
    }

    private boolean hasPipelineStage3PlanningApplied() {
        return effectiveStage3PlanningVariant() != ResultDispatchStage3Support.Stage3PlanningVariant.NONE
                && effectiveStage3PlanningVariant()
                        != ResultDispatchStage3Support.Stage3PlanningVariant.LEGACY;
    }

    private boolean showsStage3QtyMultilineDisplay() {
        return docHasActualDispatchQtyColumn() || hasPipelineStage3PlanningApplied();
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        ensureInnerTabPersistenceWired();
        shell.syncPlanInputStage2ButtonFromDispatchDirty();
        reloadSpecialRuleBadges();
        refreshDispatchPlanningStageBadge(null);
    }

    void reloadSpecialRuleBadges() {
        if (shell != null) {
            specialRuleBadges.reload(shell.snapshotUiEnv());
        }
    }

    /**
     * {@link MainShellController} の {@link javafx.stage.Stage#setOnShown} 後にフラグだけ立てる。
     * 実際の読込は {@link #onMainShellDispatchTabSelected()}（配台タブ初回選択）で行う。
     */
    void scheduleInitialReloadAfterMainWindowShown() {
        pendingInitialDispatchReload.set(true);
    }

    /** メインシェルで配台計画手動修正タブが選択されたとき（遅延ロード・保留再構築のフラッシュ）。 */
    void onMainShellDispatchTabSelected() {
        if (pendingInitialDispatchReload.compareAndSet(true, false)) {
            resetTableDisplayBeforeReload("起動時読込中");
            // タブ実体化直後の同一パルスで setGrid すると Parent.layout IOOBE になりやすい
            Platform.runLater(
                    () -> Platform.runLater(() -> reloadFromDiskQuiet(null, false, false, false)));
            return;
        }
        flushPendingGridRebuildAfterTabAttach();
    }

    private void flushPendingGridRebuildAfterTabAttach() {
        if (!pendingGridRebuildAfterTabAttach.get()) {
            return;
        }
        Platform.runLater(
                () -> {
                    if (!pendingGridRebuildAfterTabAttach.get()) {
                        return;
                    }
                    if (shell != null) {
                        shell.ensureDispatchInteractiveOnSceneForGridRebuild(false);
                    }
                    boolean onScene =
                            wideSpreadsheet != null && wideSpreadsheet.getScene() != null;
                    if (!onScene) {
                        return;
                    }
                    pendingGridRebuildAfterTabAttach.set(false);
                    rebuildGrids(this::hideReloadProgress);
                });
    }

    private void ensureDispatchColumnStripsInstalled() {
        if (wideColumnStripHost != null && wideColumnStripHost.getChildren().isEmpty()) {
            wideColumnStripHost
                    .getChildren()
                    .setAll(
                            SpreadsheetColumnSettingsStrip.create(
                                    this::resetWideColumnWidths,
                                    TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_WIDE,
                                    headerColumnCountWide,
                                    this::onWideLeadingColumnCountCommitted,
                                    null,
                                    null));
        }
        if (byDayColumnStripHost != null && byDayColumnStripHost.getChildren().isEmpty()) {
            byDayColumnStripHost
                    .getChildren()
                    .setAll(
                            SpreadsheetColumnSettingsStrip.create(
                                    this::resetByDayColumnWidths,
                                    TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_BY_DAY,
                                    headerColumnCountByDay,
                                    this::onByDayLeadingColumnCountCommitted,
                                    null,
                                    null));
        }
    }

    private int resolvedWideLeadingColumnCount() {
        return clampLeadingColumnCount(headerColumnCountWide.get(), wideSpreadsheet);
    }

    private int resolvedByDayLeadingColumnCount() {
        return clampLeadingColumnCount(headerColumnCountByDay.get(), byDaySpreadsheet);
    }

    private static int clampLeadingColumnCount(int requested, SpreadsheetView view) {
        int n = Math.max(0, requested);
        if (view != null) {
            int cols = view.getColumns().size();
            if (cols > 0) {
                n = Math.min(n, cols);
            }
        }
        return n;
    }

    private void onWideLeadingColumnCountCommitted(int n) {
        headerColumnCountWide.set(Math.max(0, n));
        SpreadsheetTabularSupport.reapplySpreadsheetColumnChrome(
                wideSpreadsheet, resolvedWideLeadingColumnCount());
    }

    private void onByDayLeadingColumnCountCommitted(int n) {
        headerColumnCountByDay.set(Math.max(0, n));
        SpreadsheetTabularSupport.reapplySpreadsheetColumnChrome(
                byDaySpreadsheet, resolvedByDayLeadingColumnCount());
    }

    private void resetWideColumnWidths() {
        if (wideSpreadsheet == null) {
            return;
        }
        double w = 112;
        for (var c : wideSpreadsheet.getColumns()) {
            c.setPrefWidth(w);
        }
    }

    private void resetByDayColumnWidths() {
        if (byDaySpreadsheet == null) {
            return;
        }
        double w = 112;
        for (var c : byDaySpreadsheet.getColumns()) {
            c.setPrefWidth(w);
        }
    }

    private void ensureDispatchColumnLayoutWatchersInstalled() {
        if (dispatchColumnLayoutWatchersInstalled) {
            return;
        }
        dispatchColumnLayoutWatchersInstalled = true;
        TableColumnOrderPersistence.installSpreadsheetColumnLayoutWatcher(
                wideSpreadsheet,
                TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_WIDE,
                () ->
                        suppressColumnReorderPersistence.get()
                                || suppressDispatchColumnLayoutPersistence.get(),
                () -> new ArrayList<>(buildWideColumnLabelsForAxis(dateAxis)));
        TableColumnOrderPersistence.installSpreadsheetColumnLayoutWatcher(
                byDaySpreadsheet,
                TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_BY_DAY,
                () ->
                        suppressColumnReorderPersistence.get()
                                || suppressDispatchColumnLayoutPersistence.get(),
                () -> new ArrayList<>(buildByDayColumnLabelsForAxis(dateAxis)));
    }

    private void applyPersistedDispatchColumnWidths(
            SpreadsheetView view,
            List<String> headers,
            TableColumnOrderPersistence.TableId tableId) {
        if (view == null || headers == null || headers.isEmpty()) {
            return;
        }
        List<TableColumnOrderPersistence.ColumnSpec> lay =
                TableColumnOrderPersistence.loadLayout(tableId);
        List<Double> widths =
                TableColumnOrderPersistence.resolveWidthsForHeaders(headers, lay, 112);
        SpreadsheetTabularSupport.applyColumnWidths(view, widths, 112);
    }

    private void ensureInnerTabPersistenceWired() {
        if (innerTabPersistenceWired || innerTabPane == null || shell == null) {
            return;
        }
        innerTabPersistenceWired = true;
        innerTabPane
                .getSelectionModel()
                .selectedIndexProperty()
                .addListener(
                        (obs, a, b) -> {
                            if (suppressInnerTabSessionPersistence.get()) {
                                return;
                            }
                            shell.persistDesktopSessionNow();
                        });
    }

    /** @return 選択中の子タブインデックス。未初期化時は -1 */
    public int snapshotInnerTabSelectedIndex() {
        if (innerTabPane == null) {
            return -1;
        }
        return innerTabPane.getSelectionModel().getSelectedIndex();
    }

    public void applyInnerTabSelectedIndex(int index) {
        if (innerTabPane == null || index < 0) {
            return;
        }
        int n = innerTabPane.getTabs().size();
        if (index >= n) {
            return;
        }
        suppressInnerTabSessionPersistence.set(true);
        try {
            innerTabPane.getSelectionModel().select(index);
        } finally {
            suppressInnerTabSessionPersistence.set(false);
        }
    }

    void clearColumnFiltersAndSort() {
        SpreadsheetTabularSupport.clearAllFiltersAndSort(wideSpreadsheet);
        SpreadsheetTabularSupport.clearAllFiltersAndSort(byDaySpreadsheet);
    }

    @FXML
    private void onClearColumnFiltersAction() {
        clearColumnFiltersAndSort();
    }

    /** 子タブに応じて「列の表示」ダイアログを開く（FXML: 列の表示）。 */
    @FXML
    private void onColumnVisibilityAction() {
        int tab = innerTabPane != null ? innerTabPane.getSelectionModel().getSelectedIndex() : 0;
        if (tab > 1) {
            return;
        }
        if (tab == 0) {
            openWideColumnVisibilityDialog();
        } else if (tab == 1) {
            openByDayColumnVisibilityDialog();
        }
    }

    private void openWideColumnVisibilityDialog() {
        Stage st = shell != null ? shell.getPrimaryStage() : null;
        ColumnVisibilitySupport.openSpreadsheetColumnVisibilityDialog(
                st,
                TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_WIDE,
                wideSpreadsheet,
                () -> new ArrayList<>(buildWideColumnLabelsForAxis(dateAxis)));
    }

    private void openByDayColumnVisibilityDialog() {
        Stage st = shell != null ? shell.getPrimaryStage() : null;
        ColumnVisibilitySupport.openSpreadsheetColumnVisibilityDialog(
                st,
                TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_BY_DAY,
                byDaySpreadsheet,
                () -> new ArrayList<>(buildByDayColumnLabelsForAxis(dateAxis)));
    }

    /**
     * 列の並べ替えダイアログ（先頭固定列の外側＝主に日付列）。ヘッダドラッグ並べ替えと同じく日付軸と JSON 保存用レイアウトを更新する。
     */
    @FXML
    private void onColumnReorderAction() {
        Stage st = shell != null ? shell.getPrimaryStage() : null;
        if (dateAxis.isEmpty()) {
            if (shell != null) {
                shell.appendLog("[dispatch-editor] 列の並べ替え: 表示する列がありません（JSON を読み込んでください）");
            }
            return;
        }
        int tab = innerTabPane != null ? innerTabPane.getSelectionModel().getSelectedIndex() : 0;
        if (tab != 0 && tab != 1) {
            if (shell != null) {
                shell.appendLog(
                        "[dispatch-editor] 列の並べ替えは「タスク×日付」「工程+機械×日」のタブでのみ利用できます。");
            }
            return;
        }
        boolean wideMode = tab == 0;
        List<String> headers =
                wideMode
                        ? new ArrayList<>(buildWideColumnLabelsForAxis(dateAxis))
                        : new ArrayList<>(buildByDayColumnLabelsForAxis(dateAxis));
        int fixed =
                wideMode ? WIDE_STATIC_HEADERS.size() : BY_DAY_STATIC_HEADERS.size();
        TableColumnOrderPersistence.TableId tid =
                wideMode
                        ? TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_WIDE
                        : TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_BY_DAY;
        boolean[] vis = TableColumnOrderPersistence.loadColumnVisibility(tid, headers.size());
        Optional<List<Integer>> choice =
                SpreadsheetColumnReorderDialog.showWithFixedLeading(st, headers, vis, fixed);
        if (choice.isEmpty()) {
            boolean anyMovableVisible = false;
            for (int i = fixed; i < headers.size(); i++) {
                if (vis == null || i >= vis.length || vis[i]) {
                    anyMovableVisible = true;
                    break;
                }
            }
            if (!anyMovableVisible && shell != null) {
                shell.appendLog(
                        "[dispatch-editor] 並べ替え対象の列がすべて非表示です。「列の表示」で日付列を表示してください。");
            }
            return;
        }
        applyDispatchInteractiveReorderPermutation(choice.get(), headers, wideMode);
    }

    private void applyDispatchInteractiveReorderPermutation(
            List<Integer> perm, List<String> headersSnapshot, boolean wideMode) {
        if (perm == null || headersSnapshot == null || perm.size() != headersSnapshot.size()) {
            return;
        }
        List<String> titleOrder = new ArrayList<>(perm.size());
        for (Integer idx : perm) {
            if (idx == null || idx < 0 || idx >= headersSnapshot.size()) {
                return;
            }
            titleOrder.add(headersSnapshot.get(idx));
        }
        TableColumnOrderPersistence.TableId tid =
                wideMode
                        ? TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_WIDE
                        : TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_BY_DAY;
        boolean[] oldVis =
                TableColumnOrderPersistence.loadColumnVisibility(tid, headersSnapshot.size());
        boolean[] newVis =
                TableColumnOrderPersistence.permuteVisibilityForLogicalReorder(
                        headersSnapshot, oldVis, titleOrder);
        TableColumnOrderPersistence.saveColumnVisibility(tid, newVis);

        List<LocalDate> computed = computeDateAxisList();
        int staticCols = wideMode ? WIDE_STATIC_HEADERS.size() : BY_DAY_STATIC_HEADERS.size();
        List<LocalDate> dates = parseDateTailAsDates(titleOrder, staticCols);
        if (dates == null || !sameMultisetLocalDate(dates, computed)) {
            return;
        }
        if (wideMode && !wideStaticPrefixMatches(titleOrder)) {
            return;
        }
        if (!wideMode && !byDayStaticPrefixMatches(titleOrder)) {
            return;
        }
        if (dates.equals(preferredDateAxisOrder)) {
            persistDispatchColumnLayouts(
                    wideMode ? titleOrder : buildWideColumnLabelsForAxis(dates),
                    wideMode ? buildByDayColumnLabelsForAxis(dates) : titleOrder);
            return;
        }
        preferredDateAxisOrder = new ArrayList<>(dates);
        persistDispatchColumnLayouts(
                wideMode ? titleOrder : buildWideColumnLabelsForAxis(dates),
                wideMode ? buildByDayColumnLabelsForAxis(dates) : titleOrder);
        suppressColumnReorderPersistence.set(true);
        try {
            rebuildGrids();
        } finally {
            suppressColumnReorderPersistence.set(false);
        }
    }

    /**
     * 段階1／2.0～3.2 実行中は段階2.1・配台試行 等を無効化する（{@link MainShellController#applyRunTabGating} から）。
     */
    void setStageRunProgressVisible(boolean stage1Running, boolean stage2Running) {
        stagePipelineBusy = stage1Running || stage2Running;
        applyDispatchTrialButtonEnabledState();
    }

    void setDeliveryCalendarReloadBlocking(boolean blocking) {
        boolean wasBlocking = deliveryCalendarReloadBlocking;
        deliveryCalendarReloadBlocking = blocking;
        applyDispatchTrialButtonEnabledState();
        if (wasBlocking && !blocking) {
            refreshAladdinPlanDisplayAfterDeliveryCalendarReload();
        }
    }

    /**
     * 納期管理ビュー再読込完了後: {@code shaped_aladdin_plan.json} はアラジン加工計画タブ反映時に更新されるが、
     * 配台表の再構築はそれより先に走ることがある。ルックアップだけ取り直して日付セルの (アラ計画) 行を再描画する。
     */
    private void refreshAladdinPlanDisplayAfterDeliveryCalendarReload() {
        if (doc == null || doc.rows().isEmpty()) {
            return;
        }
        if (shell != null) {
            shell.ensureDispatchInteractiveOnSceneForGridRebuild(false);
        }
        if (wideSpreadsheet != null && wideSpreadsheet.getScene() != null) {
            rebuildGrids();
            return;
        }
        pendingGridRebuildAfterTabAttach.set(true);
    }

    @FXML
    private void onAlignToAladdinPlanAction() {
        if (reloadInteractionDisabled
                || deliveryCalendarReloadBlocking
                || doc == null
                || docHasActualDispatchQtyColumn()) {
            return;
        }
        if (wideProfiles.isEmpty() || dateAxis.isEmpty()) {
            if (statusLabel != null) {
                statusLabel.setText("整列対象の行がありません");
            }
            return;
        }
        aladdinPlanLookup = loadAladdinPlanLookupForDisplay();
        if (aladdinPlanLookup.isEmpty()) {
            Alert a = new Alert(AlertType.WARNING);
            if (shell != null) {
                a.initOwner(shell.primaryStageForDialogs());
            }
            a.setTitle("アラジン計画に合わせる");
            a.setHeaderText("アラジン計画データがありません");
            a.setContentText(
                    "shaped_aladdin_plan.json が読み込めません。"
                            + "段階1成形結果またはアラジン計画の出力を確認してください。");
            a.showAndWait();
            return;
        }

        Alert confirm = new Alert(AlertType.CONFIRMATION);
        if (shell != null) {
            confirm.initOwner(shell.primaryStageForDialogs());
        }
        LocalDate operationDate = LocalDate.now();
        LocalDate alignFromDate = aladdinAlignFromDate(operationDate);
        boolean includesToday = alignFromDate.equals(operationDate);
        confirm.setTitle("アラジン計画に合わせる");
        confirm.setHeaderText(
                includesToday
                        ? "定常開始前のため、本日を含めアラジン計画に沿って再配分します"
                        : "翌日以降の数量をアラジン計画に沿って再配分します");
        confirm.setContentText(
                (includesToday
                                ? "操作日（"
                                        + operationDate
                                        + "）より前の暦日は変更しません。"
                                        + "本日を含む以降のみ、各タスク行の合計数量を維持したままロール単位で日付間に移動します。"
                                : "操作日（"
                                        + operationDate
                                        + "）以前の暦日は変更しません。"
                                        + "翌日以降のみ、各タスク行の合計数量を維持したままロール単位で日付間に移動します。")
                        + "\n換算数量が原反ロール長より小さい行は、表示上 20 m 等でも 1 ロール単位で移動します。");
        if (confirm.showAndWait().orElse(ButtonType.CANCEL) != ButtonType.OK) {
            return;
        }

        aladdinAlignMoveHighlights.clear();

        final int alignFromDayIndex = aladdinAlignFromDayIndexOnAxis(alignFromDate);
        if (alignFromDayIndex >= dateAxis.size()) {
            if (statusLabel != null) {
                statusLabel.setText(
                        "整列対象日（" + alignFromDate + "）以降の日付列がありません");
            }
            return;
        }

        List<String> cols = doc.columns();
        int changedRows = 0;
        int rollMoves = 0;
        int skippedRows = 0;
        for (int rowIdx = 0; rowIdx < wideProfiles.size(); rowIdx++) {
            Map<String, String> profile = wideProfiles.get(rowIdx);
            WideRow wr = rowIdx < wideRowItems.size() ? wideRowItems.get(rowIdx) : null;
            Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM unitInfo =
                    resolveRollUnitForWideRow(
                            wr != null ? wr : new WideRow(profile, dateAxis.size()));
            double unitM = unitInfo.unitM();

            double[] current = new double[dateAxis.size()];
            double[] aladdin = new double[dateAxis.size()];
            for (int j = 0; j < dateAxis.size(); j++) {
                LocalDate day = dateAxis.get(j);
                current[j] =
                        wr != null
                                ? wr.getAmount(j)
                                : ResultDispatchPivot.sumQuantityForProfileAndDateForWideMerge(
                                        doc.rows(),
                                        profile,
                                        day,
                                        ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
                aladdin[j] =
                        wr != null
                                ? aladdinPlanQtyForWideRow(wr, day)
                                : AladdinShapedPlanQtyLookup.lookup(
                                        aladdinPlanLookup,
                                        profile.get(ResultDispatchSchema.COL_MACHINE),
                                        profile.get("依頼NO"),
                                        day.format(ALADDIN_PLAN_DATE_FMT),
                                        profile.get(ResultDispatchSchema.COL_PROCESS));
            }

            AladdinSystemDispatchDisplayQty.TaskQtyContext qtyCtx =
                    taskQtyContextForWideProfile(profile);
            DispatchAladdinPlanAligner.RowResult aligned =
                    DispatchAladdinPlanAligner.alignRowFromDayIndex(
                            new DispatchAladdinPlanAligner.RowInput(
                                    current,
                                    aladdin,
                                    unitM,
                                    qtyCtx.usesConvertedQtyForAladdinDisplay()),
                            alignFromDayIndex);
            if (!aligned.changed()) {
                if (unitM <= 1e-9
                        || !Stage2PlanRowDispatchQtyMetrics.isQtyAlignedToRollUnit(
                                sumArray(current), unitM)) {
                    skippedRows++;
                }
                continue;
            }
            changedRows++;
            rollMoves += aligned.rollMoves();
            recordAladdinAlignMoveHighlights(
                    profile, current, aligned.newByDayIndex(), alignFromDayIndex);
            // 該当 profile に identity マッチする行を「全日付」一括削除してから target で再追加する。
            // upsertAllocationForWideMerge は「同日付 かつ identity マッチ」両方が必要なため、
            // 配台日カラムの表記揺れ（"2026-05-26" / "2026/05/26" / "2026-5-26" 等）が混在すると
            // 移動元の行（aladdin が 0 の日）が sameDate=false で削除されず、
            // (段階2後) として旧数値が残ってしまう。
            final Map<String, String> profileForRemoval = profile;
            doc.rows()
                    .removeIf(
                            row ->
                                    ResultDispatchPivot.matchesWideMergeIdentity(
                                            profileForRemoval,
                                            row,
                                            ResultDispatchPivot
                                                    .DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS));
            for (int j = 0; j < dateAxis.size(); j++) {
                double q = aligned.newByDayIndex()[j];
                if (q <= 1e-9) {
                    continue;
                }
                LinkedHashMap<String, String> neo = new LinkedHashMap<>();
                String dateStr = dateAxis.get(j).toString();
                for (String col : cols) {
                    if (col.equals(ResultDispatchSchema.COL_DISPATCH_DATE)) {
                        neo.put(col, dateStr);
                    } else if (col.equals(ResultDispatchSchema.COL_DISPATCH_QTY)) {
                        neo.put(col, ResultDispatchNormalizer.formatQty(q));
                    } else {
                        neo.put(col, profile.getOrDefault(col, ""));
                    }
                }
                doc.rows().add(neo);
            }
        }
        ResultDispatchNormalizer.normalizeInPlace(cols, doc.rows());
        if (changedRows > 0) {
            captureStage3TrialPlanQtySnapshotFromDocument(doc, dateAxis);
        }
        rebuildGrids();
        markDispatchDocDirty();
        if (statusLabel != null) {
            if (changedRows <= 0) {
                statusLabel.setText(
                        skippedRows > 0
                                ? "アラジン計画に合わせる対象がありませんでした（"
                                        + skippedRows
                                        + " 行スキップ）"
                                : "変更はありませんでした");
            } else {
                statusLabel.setText(
                        "アラジン計画に合わせました: "
                                + changedRows
                                + " 行・約 "
                                + rollMoves
                                + " ロール移動（黄=移動元・青=移動先）");
            }
        }
    }

    private void clearAladdinAlignMoveHighlights() {
        aladdinAlignMoveHighlights.clear();
    }

    private static String aladdinAlignMoveCellKey(Map<String, String> profile, LocalDate day) {
        if (profile == null || day == null) {
            return "";
        }
        return DispatchTrialShortages.wideShortfallKey(
                profile.get("依頼NO"),
                profile.get(ResultDispatchSchema.COL_MACHINE),
                day.toString());
    }

    private void recordAladdinAlignMoveHighlights(
            Map<String, String> profile,
            double[] before,
            double[] after,
            int fromDayIndex) {
        if (profile == null || before == null || after == null || dateAxis == null) {
            return;
        }
        int n = Math.min(before.length, Math.min(after.length, dateAxis.size()));
        for (int j = Math.max(0, fromDayIndex); j < n; j++) {
            double b = Math.max(0.0, before[j]);
            double a = Math.max(0.0, after[j]);
            if (Math.abs(b - a) <= 1e-3) {
                continue;
            }
            String key = aladdinAlignMoveCellKey(profile, dateAxis.get(j));
            if (key.isEmpty()) {
                continue;
            }
            if (a > b + 1e-3) {
                aladdinAlignMoveHighlights.put(key, AladdinAlignMoveHighlight.DEST);
            } else if (a < b - 1e-3) {
                aladdinAlignMoveHighlights.put(key, AladdinAlignMoveHighlight.SOURCE);
            }
        }
    }

    private AladdinAlignMoveHighlight aladdinAlignMoveHighlightForCell(
            Map<String, String> profile, LocalDate day) {
        if (profile == null || day == null || aladdinAlignMoveHighlights.isEmpty()) {
            return null;
        }
        return aladdinAlignMoveHighlights.get(aladdinAlignMoveCellKey(profile, day));
    }

    private AladdinAlignMoveHighlight aladdinAlignMoveHighlightForWideRow(WideRow wr, int dateIdx) {
        if (wr == null || dateIdx < 0 || dateIdx >= dateAxis.size()) {
            return null;
        }
        return aladdinAlignMoveHighlightForCell(wr.profileMap(), dateAxis.get(dateIdx));
    }

    private AladdinAlignMoveHighlight aladdinAlignMoveHighlightForByDay(ByDayRow br, int dateIdx) {
        if (br == null || dateIdx < 0 || dateIdx >= dateAxis.size() || aladdinAlignMoveHighlights.isEmpty()) {
            return null;
        }
        LocalDate day = dateAxis.get(dateIdx);
        String process = br.process() != null ? br.process().strip() : "";
        String machine = br.machine() != null ? br.machine().strip() : "";
        boolean anySource = false;
        for (Map<String, String> profile : wideProfiles) {
            String p = profile.getOrDefault(ResultDispatchSchema.COL_PROCESS, "").strip();
            String m = profile.getOrDefault(ResultDispatchSchema.COL_MACHINE, "").strip();
            if (!process.equals(p) || !machine.equals(m)) {
                continue;
            }
            AladdinAlignMoveHighlight h = aladdinAlignMoveHighlightForCell(profile, day);
            if (h == AladdinAlignMoveHighlight.DEST) {
                return AladdinAlignMoveHighlight.DEST;
            }
            if (h == AladdinAlignMoveHighlight.SOURCE) {
                anySource = true;
            }
        }
        return anySource ? AladdinAlignMoveHighlight.SOURCE : null;
    }

    private static String aladdinAlignMoveCellStyle(AladdinAlignMoveHighlight move) {
        if (move == AladdinAlignMoveHighlight.SOURCE) {
            return DATE_CELL_STYLE_ALADDIN_ALIGN_SOURCE;
        }
        if (move == AladdinAlignMoveHighlight.DEST) {
            return DATE_CELL_STYLE_ALADDIN_ALIGN_DEST;
        }
        return null;
    }

    private static double sumArray(double[] values) {
        double sum = 0.0;
        if (values == null) {
            return sum;
        }
        for (double v : values) {
            sum += v;
        }
        return sum;
    }

    /** アラジン整列: 整列開始暦日に対応する日付軸 index（該当列が無ければ {@code dateAxis.size()}）。 */
    private int aladdinAlignFromDayIndexOnAxis(LocalDate fromDate) {
        if (dateAxis == null || dateAxis.isEmpty()) {
            return 0;
        }
        for (int j = 0; j < dateAxis.size(); j++) {
            if (!dateAxis.get(j).isBefore(fromDate)) {
                return j;
            }
        }
        return dateAxis.size();
    }

    private LocalDate aladdinAlignFromDate(LocalDate operationDate) {
        return DispatchAladdinPlanAligner.resolveAlignFromDate(
                operationDate, LocalTime.now(), resolveRegularShiftStartForAladdinAlign());
    }

    private Optional<LocalTime> resolveRegularShiftStartForAladdinAlign() {
        Map<String, String> ui = shell != null ? shell.snapshotUiEnv() : Map.of();
        Path master = AppPaths.resolveMasterWorkbookPathResolved(ui, "");
        if (master == null || !Files.isRegularFile(master)) {
            return Optional.empty();
        }
        try {
            return Stage2MasterFactoryHoursReader.readRegularShift(master).start();
        } catch (IOException ignored) {
            return Optional.empty();
        }
    }

    private AladdinSystemDispatchDisplayQty.TaskQtyContext taskQtyContextForWideProfile(
            Map<String, String> profile) {
        if (doc == null || profile == null) {
            return new AladdinSystemDispatchDisplayQty.TaskQtyContext(0.0, 0.0);
        }
        List<String> cols = doc.columns();
        List<String> row = new ArrayList<>(cols.size());
        for (String col : cols) {
            row.add(profile.getOrDefault(col, ""));
        }
        return AladdinSystemDispatchDisplayQty.contextFromDispatchRow(
                cols, row, rollUnitTablesCached());
    }

    @FXML
    private void onLoadAction() {
        if (reloadInteractionDisabled) {
            return;
        }
                resetTableDisplayBeforeReload("再読込中（表示をクリア）");
        // 同一パルスで即 reload すると、空グリッドのレイアウト前に onSucceeded が走り「クリアされない」ように見える
        Platform.runLater(() -> reloadFromDiskQuiet(null, false, true, false));
    }

    @FXML
    private void onSaveAction() {
        if (shell == null) {
            return;
        }
        if (reloadInteractionDisabled) {
            return;
        }
        Path jsonPath = AppPaths.resolveResultDispatchTableJsonPath(shell.snapshotUiEnv());
        ResultDispatchDocument toWrite = doc.copy();
        Path pyExe = resolvePythonExe();
        Path pyDir = AppPaths.resolvePythonScriptDir(shell.snapshotUiEnv());

        statusLabel.setText("保存中…");
        showReloadProgress();

        Task<DispatchSaveOutcome> task =
                new Task<>() {
                    @Override
                    protected DispatchSaveOutcome call() throws Exception {
                        ResultDispatchJsonIo.write(jsonPath, toWrite);
                        String xlsxOut =
                                ResultDispatchPythonExport.exportXlsxNearJson(jsonPath, pyExe, pyDir);
                        return new DispatchSaveOutcome(jsonPath, xlsxOut);
                    }
                };
        task.setOnSucceeded(
                e -> {
                    try {
                        DispatchSaveOutcome r = task.getValue();
                        clearDispatchDocDirty();
                        statusLabel.setText("保存しました");
                        shell.appendLog("[dispatch-editor] saved json: " + r.jsonPath());
                        if (r.xlsxStdoutLine() != null && !r.xlsxStdoutLine().isEmpty()) {
                            shell.appendLog("[dispatch-editor] xlsx: " + r.xlsxStdoutLine());
                        } else {
                            shell.appendLog("[dispatch-editor] xlsx export skipped or failed (Python)");
                        }
                        showDispatchSaveFinishedDialog(r.jsonPath(), r.xlsxStdoutLine());
                    } finally {
                        hideReloadProgress();
                    }
                });
        task.setOnFailed(
                e -> {
                    try {
                        statusLabel.setText("保存エラー");
                        Throwable ex = task.getException();
                        String detail = ex != null ? ex.getMessage() : "";
                        shell.appendLog(
                                "[dispatch-editor] save failed: "
                                        + (detail != null && !detail.isBlank() ? detail : "(不明)"));
                        String msg =
                                detail != null && !detail.isBlank()
                                        ? detail
                                        : (ex != null ? ex.getClass().getSimpleName() : "(不明)");
                        shell.showErrorDialog("保存エラー", "保存に失敗しました。\n" + msg);
                    } finally {
                        hideReloadProgress();
                    }
                });
        new Thread(task, "dispatch-editor-save").start();
    }

    /** 保存ボタン後: JSON / xlsx の結果をダイアログで通知する。 */
    private void showDispatchSaveFinishedDialog(Path jsonPath, String xlsxStdoutLine) {
        boolean xlsxOk = xlsxStdoutLine != null && !xlsxStdoutLine.isBlank();
        Alert alert = new Alert(xlsxOk ? AlertType.INFORMATION : AlertType.WARNING);
        alert.setTitle("保存");
        alert.setHeaderText(xlsxOk ? "保存が完了しました" : "JSON は保存しました（Excel に注意）");
        StringBuilder text = new StringBuilder();
        text.append("JSON を保存しました。\n").append(jsonPath);
        if (xlsxOk) {
            text.append("\n\n結果配台表の Excel (xlsx) を出力しました。\n").append(xlsxStdoutLine.trim());
        } else {
            text.append(
                    "\n\nExcel (xlsx) は出力されませんでした（スクリプト未配置・タイムアウト・終了コード異常など）。"
                            + " 実行・ログのメッセージも確認してください。");
        }
        alert.setContentText(text.toString());
        shell.prepareDialogForMainTheme(alert);
        alert.showAndWait();
    }

    private void startDispatchTrial() {
        startDispatchTrialInternal();
    }

    /**
     * 配台試行: {@link ResultDispatchTrialPython} および不足 JSON 連携（段階2同一エンジン。
     * JSON 暦日数量は配台後の未達照合・配台日スライド用。配台ループ中の上限にはしない）。
     * ログは「実行・ログ」タブへ出力する。
     */
    private void startDispatchTrialInternal() {
        if (shell == null) {
            return;
        }
        String blockLabel = "配台試行（段階3）";
        if (shell.blockIfSummaryAiDispatchExportLocked(blockLabel)) {
            return;
        }
        if (shell.blockIfMaterialLookupTablesHaveBlankValues(blockLabel)) {
            return;
        }
        if (reloadInteractionDisabled
                || deliveryCalendarReloadBlocking
                || dispatchDocDirtySinceSave) {
            return;
        }
        PipelineExecutionTimingKind timingKind = PipelineExecutionTimingKind.STAGE3;
        if (!shell.tryBeginDispatchTrialGating(timingKind)) {
            return;
        }
        statusLabel.setText("配台試行中...");
        if (dispatchTrialButton != null) {
            dispatchTrialButton.setDisable(true);
        }
        showReloadProgress();
        Path jsonPath = AppPaths.resolveResultDispatchTableJsonPath(shell.snapshotUiEnv());
        final Path trialPythonExe = resolvePythonExe();

        shell.selectMainShellTab(MainShellTabId.RUN);
        shell.appendLog("--- start: 段階3（配台試行） ---");
        shell.appendLog("[配台試行] 処理を開始しました。");
        shell.appendLog("[配台試行] Python 実行ファイル: " + trialPythonExe.toAbsolutePath().normalize());

        final ResultDispatchDocument trialInputSnapshot = doc.copy();
        captureStage3TrialPlanQtySnapshotFromDocument(
                trialInputSnapshot, snapshotDateAxisForTrialPlanQtyCapture(trialInputSnapshot));
        stage3TrialInputDocumentSnapshot = trialInputSnapshot.copy();
        Stage owner = shell.getPrimaryStage();

        Task<String> task =
                new Task<>() {
                    @Override
                    protected String call() throws Exception {
                        shell.beginPipelineExecutionTiming(timingKind);
                        try {
                            ResultDispatchDocument writeDoc = doc.copy();
                            ResultDispatchJsonIo.write(jsonPath, writeDoc);
                            shell.appendLog(
                                    "[dispatch-editor] trial: 試行前にメモリ上の表を JSON に同期 "
                                            + jsonPath.toAbsolutePath().normalize());
                            Path pyExe = trialPythonExe;
                            Path pyDir = AppPaths.resolvePythonScriptDir(shell.snapshotUiEnv());
                            Map<String, String> pyEnv = shell.snapshotDispatchTrialPythonEnv();
                            return ResultDispatchTrialPython.runTrial(
                                    jsonPath,
                                    pyExe,
                                    pyDir,
                                    pyEnv,
                                    shell::appendLog);
                        } finally {
                            shell.endPipelineExecutionTiming(timingKind);
                        }
                    }
                };

        task.setOnSucceeded(
                e -> {
                    try {
                        String shortagesPath = task.getValue();
                        statusLabel.setText("配台試行完了");
                        shell.refreshRunTabStage2ArtifactLinks();
                        shell.appendLog("[dispatch-editor] trial: " + shortagesPath);
                        shell.appendLog("[配台試行] 正常終了しました。");
                        shell.appendLog("不足情報JSON: " + shortagesPath);
                        reloadFromDiskQuietAfterDispatchTrial(
                                () -> {
                                    try {
                                        shell.reloadDeliveryCalendarInBackgroundAfterDispatchTrialSuccess();
                                        shell.notifyStage3DispatchTrialSuccess();
                                    } catch (Throwable upex) {
                                        String em =
                                                upex.getMessage() != null
                                                        ? upex.getMessage()
                                                        : upex.getClass().getSimpleName();
                                        shell.appendLog("[配台試行] 試行後処理で例外: " + em);
                                        shell.appendLog("[dispatch-editor] trial post-run: " + em);
                                        shell.notifyStage3DispatchTrialFailure(em);
                                        shell.endDispatchTrialGating(timingKind);
                                        return;
                                    }
                                    Platform.runLater(
                                            () -> {
                                                try {
                                                    showDispatchQtyShortfallDialogIfNeeded(owner);
                                                    showDispatchShortageHintsDialogIfNeeded(owner);
                                                    showRawInputMorningDispatchRateWarningIfNeeded(
                                                            owner);
                                                    DispatchTrialConsistency.CheckResult cr =
                                                            DispatchTrialConsistency.compareDocuments(
                                                                    trialInputSnapshot, doc);
                                                    if (cr.consistent()) {
                                                        shell.appendLog(
                                                                "[整合性] 保存済み表と試行後の成果物（結果_配台表.json）は、"
                                                                        + "依頼NO×機械名の当日配台数量合計および配台試行順番（工程別最小値）の観点で一致しました。");
                                                        shell.appendLog(
                                                                "[dispatch-editor] trial: 整合性OK（保存表と再読込JSONの数量・試行順）");
                                                    } else {
                                                        shell.appendLog(
                                                                "[整合性] 保存済み表と試行後の成果物に差異があります（詳細は下記）:");
                                                        for (String dl : cr.detailLines()) {
                                                            shell.appendLog(dl);
                                                        }
                                                        shell.appendLog(
                                                                "[dispatch-editor] trial: 整合性に差異あり（"
                                                                        + cr.detailLines().size()
                                                                        + " 件）— 実行・ログ参照");
                                                    }
                                                    DispatchTrialUnassignedWizard.showIfNeeded(
                                                            owner, shell, Path.of(shortagesPath));
                                                } catch (Throwable deferredEx) {
                                                    String em =
                                                            deferredEx.getMessage() != null
                                                                    ? deferredEx.getMessage()
                                                                    : deferredEx.getClass()
                                                                            .getSimpleName();
                                                    shell.appendLog(
                                                            "[配台試行] 試行後確認ダイアログで例外: " + em);
                                                    shell.appendLog(
                                                            "[dispatch-editor] trial post-run deferred: "
                                                                    + em);
                                                } finally {
                                                    shell.endDispatchTrialGating(timingKind);
                                                }
                                            });
                                });
                    } catch (Throwable sucEx) {
                        String em =
                                sucEx.getMessage() != null
                                        ? sucEx.getMessage()
                                        : sucEx.getClass().getSimpleName();
                        shell.appendLog("[配台試行] 成功ハンドラ内例外: " + em);
                        shell.appendLog("[dispatch-editor] trial onSucceeded: " + em);
                        shell.notifyStage3DispatchTrialFailure(em);
                        hideReloadProgress();
                        shell.endDispatchTrialGating(timingKind);
                    } finally {
                        Platform.runLater(this::applyDispatchTrialButtonEnabledState);
                    }
                });
        task.setOnFailed(
                e -> {
                    try {
                        Throwable ex = task.getException();
                        statusLabel.setText("配台試行エラー");
                        String msg = ex != null ? ex.getMessage() : "(不明)";
                        shell.appendLog("[dispatch-editor] trial failed: " + msg);
                        shell.appendLog("[配台試行] エラーで終了しました。");
                        shell.appendLog(msg);
                        if (ex != null) {
                            java.io.StringWriter sw = new java.io.StringWriter();
                            ex.printStackTrace(new java.io.PrintWriter(sw));
                            String stack = sw.toString();
                            int max = 8000;
                            if (stack.length() > max) {
                                stack = stack.substring(0, max) + "\n... (truncated)";
                            }
                            for (String ln : stack.split("\n")) {
                                if (!ln.isEmpty()) {
                                    shell.appendLog(ln);
                                }
                            }
                        }
                        shell.notifyStage3DispatchTrialFailure(msg);
                    } catch (Throwable handlerEx) {
                        shell.appendLog(
                                "[dispatch-editor] trial onFailed handler: "
                                        + handlerEx.getMessage());
                    } finally {
                        hideReloadProgress();
                        shell.endDispatchTrialGating(timingKind);
                        Platform.runLater(this::applyDispatchTrialButtonEnabledState);
                    }
                });
        task.setOnCancelled(
                e -> {
                    try {
                        statusLabel.setText("配台試行キャンセル");
                        shell.appendLog("[配台試行] キャンセルされました。");
                    } finally {
                        hideReloadProgress();
                        shell.endDispatchTrialGating(timingKind);
                        Platform.runLater(this::applyDispatchTrialButtonEnabledState);
                    }
                });
        new Thread(task, "dispatch-trial").start();
    }

    @FXML
    private void onDispatchTrialAction() {
        startDispatchTrial();
    }

    @FXML
    private void onBuildStage3InputAction() {
        if (shell == null || stage3InputBuildBusy) {
            return;
        }
        shell.triggerBuildStage3Input();
    }

    /** 入力3表生成の実行中フラグ（二重起動防止・UI 連動用）。 */
    boolean isStage3InputBuildBusy() {
        return stage3InputBuildBusy;
    }

    /**
     * 入力3表生成中のプログレス表示とボタン状態を切り替える。
     * ツールバー右の {@link #reloadProgressBar} / {@link #busyIndicator} を不定プログレスで表示する。
     */
    void setStage3InputBuildProgressVisible(boolean visible) {
        Runnable apply =
                () -> {
                    stage3InputBuildBusy = visible;
                    if (reloadProgressBar != null) {
                        reloadProgressBar.setManaged(visible);
                        reloadProgressBar.setVisible(visible);
                        reloadProgressBar.setProgress(
                                visible ? ProgressBar.INDETERMINATE_PROGRESS : 0);
                    }
                    if (busyIndicator != null) {
                        busyIndicator.setManaged(visible);
                        busyIndicator.setVisible(visible);
                    }
                    if (statusLabel != null) {
                        if (visible) {
                            statusLabel.setText(BUILD_STAGE3_INPUT_STATUS_BUSY);
                        } else if (BUILD_STAGE3_INPUT_STATUS_BUSY.equals(statusLabel.getText())) {
                            statusLabel.setText("");
                        }
                    }
                    applyDispatchTrialButtonEnabledState();
                };
        if (Platform.isFxApplicationThread()) {
            apply.run();
        } else {
            Platform.runLater(apply);
        }
    }

    @FXML
    private void onWideRowUpAction() {
        int i = selectedWideProfileIndex();
        if (i <= 0) {
            return;
        }
        int colIdx = wideSpreadsheetFocusedColumnIndex();
        swapWideProfiles(i - 1, i);
        focusWideProfileCellAfterReorder(i - 1, colIdx);
    }

    @FXML
    private void onWideRowDownAction() {
        int i = selectedWideProfileIndex();
        if (i < 0 || i >= wideProfiles.size() - 1) {
            return;
        }
        int colIdx = wideSpreadsheetFocusedColumnIndex();
        swapWideProfiles(i, i + 1);
        focusWideProfileCellAfterReorder(i + 1, colIdx);
    }

    private void reloadFromDiskQuiet() {
        reloadFromDiskQuiet(null, false, false, false);
    }

    /**
     * 配台試行（段階3）正常終了後のみ: {@code dispatch_trial_shortages.json} を結果 JSON と同階層から読み、未達表・赤セル
     * ヒントを同期する。通常の「再読み」や段階2後の再同期では {@link #reloadFromDiskQuiet()} を使い、古い不足 JSON
     * を読まない。
     */
    private void reloadFromDiskQuietAfterDispatchTrial(Runnable afterSuccessOnFxThread) {
        resetTableDisplayBeforeReload("配台試行後（再読込中）");
        Platform.runLater(() -> reloadFromDiskQuiet(afterSuccessOnFxThread, true, false, false));
    }

    /** 子タブが遅延プレースホルダのとき実 Spreadsheet をシーンに戻す。 */
    void ensureInnerTabsMaterializedForRebuild() {
        if (innerTabPane == null || shell == null) {
            return;
        }
        for (Tab t : innerTabPane.getTabs()) {
            shell.restoreDeferredTabContent(t);
        }
    }

    /**
     * Reloads JSON from disk asynchronously; runs {@code afterSuccessOnFxThread} on the FX thread after grids are
     * rebuilt (only when load succeeds).
     *
     * @param applyDispatchTrialShortfallJson true のときのみ隣接の {@code dispatch_trial_shortages.json} を読む（配台試行
     *     直後）。false のときは未達表・セル赤表示用キーをクリアする（再読み・外部更新後など、メイン JSON と不足 JSON
     *     の生成タイミングがずれていると誤表示になるため）。
     * @param userCompletionDialog ユーザーが「再読み」ボタンを押したときのみ true（自動再同期ではダイアログを出さない）。
     * @param validatePlanInputCoverage 段階2直後など、タスク入力と手動修正表の整合を検証して漏れがあればエラーダイアログを出す。
     */
    private void reloadFromDiskQuiet(
            Runnable afterSuccessOnFxThread,
            boolean applyDispatchTrialShortfallJson,
            boolean userCompletionDialog,
            boolean validatePlanInputCoverage) {
        reloadFromDiskQuiet(
                afterSuccessOnFxThread,
                applyDispatchTrialShortfallJson,
                userCompletionDialog,
                validatePlanInputCoverage,
                null);
    }

    private void reloadFromDiskQuiet(
            Runnable afterSuccessOnFxThread,
            boolean applyDispatchTrialShortfallJson,
            boolean userCompletionDialog,
            boolean validatePlanInputCoverage,
            Path jsonOverride) {
        if (shell == null) {
            return;
        }
        Path p =
                jsonOverride != null
                        ? jsonOverride
                        : AppPaths.resolveResultDispatchTableJsonPath(shell.snapshotUiEnv());
        jsonPathLabel.setText(p.toString());
        bumpDispatchSpreadsheetLayoutGeneration();
        if (!Files.isRegularFile(p)) {
            statusLabel.setText("ファイルなし");
            doc = ResultDispatchDocument.empty();
            timelineCalendarMeters = DispatchTimelineCalendarMetersIndex.empty();
            clearAladdinAlignMoveHighlights();
            clearDispatchShortfallUi();
            rebuildGrids(this::hideReloadProgress);
            clearDispatchDocDirty();
            if (applyDispatchTrialShortfallJson && shell != null) {
                shell.endActiveDispatchTrialGatingIfAny();
            }
            if (validatePlanInputCoverage) {
                showPlanInputCoverageGapErrorIfNeeded(p);
            }
            if (userCompletionDialog) {
                shell.showWarningDialog("再読み", "結果_配台表.json が見つかりません。\n" + p);
            }
            return;
        }
        SpreadsheetTabularSupport.showScratchGridWhileReloading(wideSpreadsheet);
        SpreadsheetTabularSupport.showScratchGridWhileReloading(byDaySpreadsheet);
        showReloadProgress();
        final Path jsonPath = p;
        Task<ReloadBundle> task =
                new Task<>() {
                    @Override
                    protected ReloadBundle call() throws Exception {
                        ResultDispatchDocument d = ResultDispatchJsonIo.read(jsonPath);
                        DispatchTimelineCalendarMetersIndex timeline =
                                DispatchTimelineCalendarMetersIndex.tryLoadNearResultDispatchJson(
                                        jsonPath);
                        return new ReloadBundle(d, timeline);
                    }
                };
        task.setOnSucceeded(
                ev -> {
                    ReloadBundle b = task.getValue();
                    doc = b.doc();
                    timelineCalendarMeters = b.timelineMeters();
                    clearAladdinAlignMoveHighlights();
                    if (validatePlanInputCoverage) {
                        showPlanInputCoverageGapErrorIfNeeded(jsonPath);
                    }
                    boolean stage2ColsFilled =
                            ResultDispatchStage2ColumnSupport.ensureStage2RequiredColumns(doc);
                    if (stage2ColsFilled) {
                        markDispatchDocDirty();
                        statusLabel.setText(
                                doc.rows().size() + " 行（段階2必須列を補完しました。保存してください）");
                    } else {
                        statusLabel.setText(doc.rows().size() + " 行");
                    }
                    if (applyDispatchTrialShortfallJson) {
                        applyDispatchShortfallFromDisk(jsonPath);
                        if (stage3TrialPlanQtySnapshot.isEmpty()) {
                            scheduleStage3TrialPlanQtySnapshotCapture();
                        }
                        loadStage21TrialSnapshotFromDiskIfNeeded(jsonPath);
                    } else {
                        clearDispatchShortfallUi();
                        if (!retainStage21TrialMetaOnNextReload) {
                            clearStage21TrialSnapshot(jsonPath);
                        } else {
                            retainStage21TrialMetaOnNextReload = false;
                            loadStage21TrialSnapshotFromDiskIfNeeded(jsonPath);
                            loadStage21TimelineFromDisk();
                        }
                        if (!retainStage3TrialSnapshotOnNextReload) {
                            scheduleStage3TrialPlanQtySnapshotCapture();
                        } else {
                            retainStage3TrialSnapshotOnNextReload = false;
                        }
                        loadStage21TrialSnapshotFromDiskIfNeeded(jsonPath);
                    }
                    loadStage3BaselineFromSidecarIfNeeded(jsonPath);
                    refreshDispatchPlanningStageBadge(jsonPath);
                    shell.refreshLearnedSpeedDataQuietly();
                    clearDispatchDocDirty();
                    if (shell != null) {
                        shell.ensureDispatchInteractiveOnSceneForGridRebuild(userCompletionDialog);
                    }
                    boolean wideOnScene =
                            wideSpreadsheet != null && wideSpreadsheet.getScene() != null;
                    Runnable afterLayouts =
                            buildReloadSuccessAfterLayoutsRunnable(
                                    afterSuccessOnFxThread,
                                    userCompletionDialog,
                                    stage2ColsFilled,
                                    jsonPath);
                    if (!wideOnScene) {
                        pendingGridRebuildAfterTabAttach.set(true);
                        hideReloadProgress();
                        if (afterSuccessOnFxThread != null) {
                            afterSuccessOnFxThread.run();
                        }
                        if (userCompletionDialog) {
                            shell.showInformationDialog(
                                    "再読み完了",
                                    doc.rows().size()
                                            + " 行を読み込みました（配台タブを開くと表を表示します）。\n"
                                            + jsonPath);
                        }
                        return;
                    }
                    pendingGridRebuildAfterTabAttach.set(false);
                    if (afterLayouts != null) {
                        rebuildGrids(afterLayouts);
                    } else {
                        rebuildGrids(this::hideReloadProgress);
                    }
                });
        task.setOnFailed(
                ev -> {
                    Throwable loadEx = task.getException();
                    doc = ResultDispatchDocument.empty();
                    timelineCalendarMeters = DispatchTimelineCalendarMetersIndex.empty();
                    statusLabel.setText("読込エラー");
                    shell.appendLog(
                            "[dispatch-editor] load failed: "
                                    + (loadEx != null ? loadEx.getMessage() : ""));
                    clearDispatchShortfallUi();
                    rebuildGrids(this::hideReloadProgress);
                    clearDispatchDocDirty();
                    if (applyDispatchTrialShortfallJson && shell != null) {
                        shell.endActiveDispatchTrialGatingIfAny();
                    }
                    if (userCompletionDialog) {
                        String msg =
                                loadEx != null
                                                && loadEx.getMessage() != null
                                                && !loadEx.getMessage().isBlank()
                                        ? loadEx.getMessage()
                                        : (loadEx != null ? loadEx.toString() : "不明");
                        shell.showErrorDialog("読込エラー", msg);
                    }
                });
        new Thread(task, "dispatch-editor-reload").start();
    }

    /** 段階2 正常終了直後: 表を JSON から再構築する。 */
    void reloadTableFromDiskAfterStage2Success() {
        reloadTableFromDiskAfterStage2Success(null);
    }

    /**
     * 段階2 正常終了直後: 表を JSON から再構築し、成功後に {@code afterSuccessOnFxThread} を FX スレッドで実行する。
     */
    void reloadTableFromDiskAfterStage2Success(Runnable afterSuccessOnFxThread) {
        Path jsonPath = AppPaths.resolveResultDispatchTableStage2JsonPath(shell.snapshotUiEnv());
        clearStage3PlanningMeta(jsonPath);
        reloadFromDiskQuiet(afterSuccessOnFxThread, false, false, true, jsonPath);
    }

    /** 段階2.1 正常終了直後: 正本反映済み。比較 baseline を保持してメイン JSON を再読込する。 */
    void reloadTableFromDiskAfterStage21PromotedSuccess(Runnable afterSuccessOnFxThread) {
        retainStage21TrialMetaOnNextReload = true;
        reloadFromDiskQuiet(afterSuccessOnFxThread, false, false, false);
    }

    /** @deprecated 試行のみ（正本未反映）の旧フロー。段階2.1 成功時は {@link #reloadTableFromDiskAfterStage21PromotedSuccess} を使用。 */
    @Deprecated
    void reloadTableFromDiskAfterStage21Success(Runnable afterSuccessOnFxThread) {
        retainStage21TrialMetaOnNextReload = true;
        reloadFromDiskQuiet(afterSuccessOnFxThread, false, false, false);
    }

    /** 段階2.1 実行開始前: 段階2 baseline を sidecar へ保存する。 */
    void captureStage21BaselineBeforeRun(Path mainJsonPath, Path overtimeOverridesJson) {
        if (doc == null || dateAxis == null || dateAxis.isEmpty()) {
            return;
        }
        captureStage21BaselineFromDocument(doc, dateAxis, mainJsonPath, overtimeOverridesJson, null);
    }

    /**
     * 段階2.1 正本反映後: (段階2後)/(段階2.1後) 比較用 sidecar・タイムライン・バッジを更新する。
     */
    void finalizeStage21PromotedWithComparisonAfterRunSuccess(
            Path mainJsonPath, Path stage21ResultJson, Path overtimeOverridesJson) {
        Stage21TrialSnapshotStore.OverrideSummary summary =
                OvertimeSimulationOverridesReader.summarize(overtimeOverridesJson);
        if (stage21BaselinePlanQtySnapshot.isEmpty() && mainJsonPath != null) {
            stage21BaselinePlanQtySnapshot.putAll(
                    Stage21TrialSnapshotStore.tryLoadMeta(mainJsonPath).entries());
        }
        if (mainJsonPath != null) {
            Stage21TrialSnapshotStore.writePromotedWithComparison(
                    mainJsonPath,
                    stage21BaselinePlanQtySnapshot,
                    stage21ResultJson,
                    overtimeOverridesJson,
                    summary);
        }
        stage21TrialApplied = !stage21BaselinePlanQtySnapshot.isEmpty();
        stage21TrialMeta =
                mainJsonPath != null
                        ? Stage21TrialSnapshotStore.tryLoadMeta(mainJsonPath)
                        : Stage21TrialSnapshotStore.Stage21TrialMeta.empty();
        loadStage21TimelineFromDisk();
        refreshStage21AttendanceApplyPanel(mainJsonPath);
        refreshDispatchPlanningStageBadge(mainJsonPath);
        if (statusLabel != null && stage21TrialApplied) {
            statusLabel.setText(formatStage21CompletionStatus(stage21TrialMeta));
        }
    }

    /** @deprecated {@link #finalizeStage21PromotedWithComparisonAfterRunSuccess} を使用。 */
    @Deprecated
    void applyStage21PromotedMetaAfterRunSuccess(Path mainJsonPath) {
        finalizeStage21PromotedWithComparisonAfterRunSuccess(mainJsonPath, null, null);
    }

    /** @deprecated 試行のみ（正本未反映）の旧フロー。 */
    @Deprecated
    void finalizeStage21TrialAfterRunSuccess(
            Path mainJsonPath, Path stage21ResultJson, Path overtimeOverridesJson) {
        Stage21TrialSnapshotStore.OverrideSummary summary =
                OvertimeSimulationOverridesReader.summarize(overtimeOverridesJson);
        stage21TrialApplied = true;
        stage21TrialMeta =
                new Stage21TrialSnapshotStore.Stage21TrialMeta(
                        true,
                        false,
                        stage21ResultJson != null
                                ? stage21ResultJson.toAbsolutePath().normalize().toString()
                                : "",
                        overtimeOverridesJson != null
                                ? overtimeOverridesJson.toAbsolutePath().normalize().toString()
                                : "",
                        summary,
                        java.time.LocalDateTime.now()
                                .format(java.time.format.DateTimeFormatter.ISO_LOCAL_DATE_TIME),
                        Map.copyOf(stage21BaselinePlanQtySnapshot));
        if (mainJsonPath != null) {
            Stage21TrialSnapshotStore.writeWithMeta(
                    mainJsonPath,
                    stage21BaselinePlanQtySnapshot,
                    stage21ResultJson,
                    overtimeOverridesJson,
                    summary);
        }
        loadStage21TimelineFromDisk();
        refreshStage21AttendanceApplyPanel(mainJsonPath);
    }

    private void loadStage21TimelineFromDisk() {
        if (shell == null) {
            stage21TimelineCalendarMeters = DispatchTimelineCalendarMetersIndex.empty();
            return;
        }
        Path stage21Json = AppPaths.resolveStage21ResultDispatchJsonPath(shell.snapshotUiEnv());
        Path fromMeta =
                stage21TrialMeta != null ? stage21TrialMeta.stage21ResultDispatchPath() : null;
        if (fromMeta != null && Files.isRegularFile(fromMeta)) {
            stage21Json = fromMeta;
        }
        stage21TimelineCalendarMeters =
                DispatchTimelineCalendarMetersIndex.tryLoadNearResultDispatchJson(stage21Json);
        if (!stage21TimelineCalendarMeters.isLoaded()) {
            Path mainJson = AppPaths.resolveResultDispatchTableStage2JsonPath(shell.snapshotUiEnv());
            if (Files.isRegularFile(mainJson)) {
                stage21TimelineCalendarMeters =
                        DispatchTimelineCalendarMetersIndex.tryLoadNearResultDispatchJson(
                                mainJson);
            }
        }
        if (!stage21TimelineCalendarMeters.isLoaded() && shell != null) {
            shell.appendLog(
                    "[stage2.1] 警告: 設備ガント契約が読めないため (段階2.1後) は 0 表示になります: "
                            + stage21Json);
        }
    }

    /** 再読込や段階2開始の直前に、メモリ上の表を空表示へ戻す（JSON パスラベルは維持）。 */
    private void resetTableDisplayBeforeReload(String statusText) {
        if (shell == null) {
            return;
        }
        clearColumnFiltersAndSort();
        preferredDateAxisOrder = null;
        Path p = AppPaths.resolveResultDispatchTableJsonPath(shell.snapshotUiEnv());
        if (jsonPathLabel != null) {
            jsonPathLabel.setText(p.toString());
        }
        bumpDispatchSpreadsheetLayoutGeneration();
        doc = ResultDispatchDocument.empty();
        timelineCalendarMeters = DispatchTimelineCalendarMetersIndex.empty();
        clearDispatchShortfallUi();
        SpreadsheetTabularSupport.showScratchGridWhileReloading(wideSpreadsheet);
        SpreadsheetTabularSupport.showScratchGridWhileReloading(byDaySpreadsheet);
        clearDispatchDocDirty();
        if (statusLabel != null) {
            statusLabel.setText(statusText);
        }
        if (wideSpreadsheet != null && wideSpreadsheet.getScene() != null) {
            wideSpreadsheet.requestLayout();
        }
        if (byDaySpreadsheet != null && byDaySpreadsheet.getScene() != null) {
            byDaySpreadsheet.requestLayout();
        }
    }

    /**
     * 段階2パイプライン開始時に、古い行を誤認しないよう表の表示を空にする（JSON パスラベルは維持）。
     */
    void resetTableDisplayForStage2Run() {
        if (shell != null) {
            clearStage3PlanningMeta(
                    AppPaths.resolveResultDispatchTableJsonPath(shell.snapshotUiEnv()));
        }
        resetTableDisplayBeforeReload("段階2実行中（表示をクリア）");
    }

    /** 段階3.0/3.1/3.2 正常終了直後: baseline を保持したまま JSON を再読込する。 */
    void reloadTableFromDiskAfterStage3PipelineSuccess() {
        reloadTableFromDiskAfterStage3PipelineSuccess(null);
    }

    void reloadTableFromDiskAfterStage3PipelineSuccess(Runnable afterSuccessOnFxThread) {
        retainStage3TrialSnapshotOnNextReload = true;
        reloadFromDiskQuiet(afterSuccessOnFxThread, false, false, false);
    }

    /**
     * 段階3.0/3.1/3.2 実行直前: (段階3前) baseline をメモリと sidecar へ保存する。
     * 実行・ログタブから起動するときも正しい baseline にするため、可能ならディスク上の JSON から読む。
     */
    void persistStage3BaselineBeforePipelineRun() {
        if (shell == null) {
            return;
        }
        Path jsonPath = AppPaths.resolveResultDispatchTableJsonPath(shell.snapshotUiEnv());
        try {
            ResultDispatchDocument sourceDoc =
                    Files.isRegularFile(jsonPath)
                            ? ResultDispatchJsonIo.read(jsonPath)
                            : doc;
            List<LocalDate> axis = snapshotDateAxisForTrialPlanQtyCapture(sourceDoc);
            List<Map<String, String>> profiles =
                    ResultDispatchPivot.distinctWideTaskProfiles(
                            sourceDoc.columns(),
                            sourceDoc.rows(),
                            ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
            captureStage3TrialPlanQtySnapshotFromDocument(sourceDoc, profiles, axis);
            Stage3PlanningMetaStore.writeBaselineEntries(jsonPath, stage3TrialPlanQtySnapshot);
        } catch (Exception ex) {
            shell.appendLog(
                    "[stage3] baseline 保存失敗: "
                            + (ex.getMessage() != null ? ex.getMessage() : ex));
        }
    }

    private void loadStage3BaselineFromSidecarIfNeeded(Path dispatchJsonPath) {
        if (dispatchJsonPath == null || !stage3TrialPlanQtySnapshot.isEmpty()) {
            return;
        }
        Map<String, Double> baseline = Stage3PlanningMetaStore.readBaselineEntries(dispatchJsonPath);
        if (!baseline.isEmpty()) {
            stage3TrialPlanQtySnapshot.clear();
            stage3TrialPlanQtySnapshot.putAll(baseline);
        }
    }

    /**
     * 外部で {@code 結果_配台表.json} が更新されたあと、当タブの表をディスクから再読込する（段階2終了後の再同期やワークスペース復元後など）。
     */
    void reloadTableFromDiskAfterExternalUpdate() {
        reloadFromDiskQuiet(null, false, false, false);
    }

    /** 配台ワークスペース用スナップショット: メモリ上の配台表ドキュメントのコピー（UI スレッド）。 */
    public ResultDispatchDocument copyDispatchDocumentForSnapshot() {
        return doc.copy();
    }

    private void showReloadProgress() {
        if (reloadProgressBar != null) {
            reloadProgressBar.setManaged(true);
            reloadProgressBar.setVisible(true);
            reloadProgressBar.setProgress(ProgressBar.INDETERMINATE_PROGRESS);
        }
        if (busyIndicator != null) {
            busyIndicator.setManaged(true);
            busyIndicator.setVisible(true);
        }
        setReloadInteractionDisabled(true);
    }

    private void hideReloadProgress() {
        if (reloadProgressBar != null) {
            reloadProgressBar.setProgress(0);
            reloadProgressBar.setVisible(false);
            reloadProgressBar.setManaged(false);
        }
        if (busyIndicator != null) {
            busyIndicator.setVisible(false);
            busyIndicator.setManaged(false);
        }
        setReloadInteractionDisabled(false);
    }

    private void setReloadInteractionDisabled(boolean disabled) {
        reloadInteractionDisabled = disabled;
        if (loadButton != null) {
            loadButton.setDisable(disabled);
        }
        if (saveButton != null) {
            saveButton.setDisable(disabled);
        }
        applyDispatchTrialButtonEnabledState();
        if (wideRowUpButton != null) {
            wideRowUpButton.setDisable(disabled);
        }
        if (wideRowDownButton != null) {
            wideRowDownButton.setDisable(disabled);
        }
        applyAlignToAladdinPlanButtonEnabledState();
    }

    private void markDispatchDocDirty() {
        dispatchDocDirtySinceSave = true;
        applyDispatchTrialButtonEnabledState();
        if (shell != null) {
            shell.onDispatchInteractiveTableDirtyChanged(true);
        }
    }

    private void clearDispatchDocDirty() {
        dispatchDocDirtySinceSave = false;
        applyDispatchTrialButtonEnabledState();
        if (shell != null) {
            shell.onDispatchInteractiveTableDirtyChanged(false);
        }
    }

    /** 最後の保存または「再読み」以降に手動編集があり、ディスク上の JSON と表が一致しないとき true。 */
    boolean isDispatchDocDirtySinceSave() {
        return dispatchDocDirtySinceSave;
    }

    /**
     * 配台試行ボタン: 再読込中は無効。表を手動編集して未保存のときは無効（保存または「再読み」で有効化）。
     */
    void refreshSummaryExportLockPresentation() {
        applyDispatchTrialButtonEnabledState();
    }

    private boolean isSummaryExportLockedByLockFile() {
        return shell != null && shell.isSummaryAiDispatchExportLocked();
    }

    private void applyDispatchTrialButtonEnabledState() {
        boolean blockTrial =
                reloadInteractionDisabled
                        || deliveryCalendarReloadBlocking
                        || dispatchDocDirtySinceSave
                        || isSummaryExportLockedByLockFile()
                        || stagePipelineBusy;
        if (dispatchTrialButton != null) {
            dispatchTrialButton.setDisable(blockTrial);
        }
        if (buildStage3InputButton != null) {
            boolean blockBuild3Input =
                    reloadInteractionDisabled
                            || deliveryCalendarReloadBlocking
                            || dispatchDocDirtySinceSave
                            || isSummaryExportLockedByLockFile()
                            || stagePipelineBusy
                            || stage3InputBuildBusy;
            buildStage3InputButton.setDisable(blockBuild3Input);
            buildStage3InputButton.setText(
                    stage3InputBuildBusy
                            ? BUILD_STAGE3_INPUT_BUTTON_TEXT_BUSY
                            : BUILD_STAGE3_INPUT_BUTTON_TEXT_DEFAULT);
        }
        if (isSummaryExportLockedByLockFile() && !reloadInteractionDisabled && !deliveryCalendarReloadBlocking) {
            Tooltip t =
                    new Tooltip(
                            "サマリ xlsx を作成中です。完了後に配台試行するか、実行・ログタブの「ロック解除」を使用してください。");
            if (dispatchTrialButton != null) {
                dispatchTrialButton.setTooltip(t);
            }
            if (buildStage3InputButton != null) {
                buildStage3InputButton.setTooltip(t);
            }
        } else if (deliveryCalendarReloadBlocking && !reloadInteractionDisabled) {
            Tooltip t =
                    new Tooltip("納期管理ビューを再読み込み中です。完了後に配台試行を実行してください。");
            if (dispatchTrialButton != null) {
                dispatchTrialButton.setTooltip(t);
            }
            if (buildStage3InputButton != null) {
                buildStage3InputButton.setTooltip(
                        new Tooltip("納期管理ビューを再読み込み中です。完了後に入力3表を生成してください。"));
            }
        } else if (dispatchDocDirtySinceSave && !reloadInteractionDisabled) {
            Tooltip t =
                    new Tooltip(
                            "未保存の編集があります。「保存 (JSON+xlsx)」または「再読み」で確定してから配台試行を実行してください。");
            if (dispatchTrialButton != null) {
                dispatchTrialButton.setTooltip(t);
            }
            if (buildStage3InputButton != null) {
                buildStage3InputButton.setTooltip(
                        new Tooltip(
                                "未保存の編集があります。「保存 (JSON+xlsx)」で確定してから入力3表を生成してください。"));
            }
        } else if (stage3InputBuildBusy) {
            Tooltip t = new Tooltip("入力3表を生成しています。完了までお待ちください。");
            if (buildStage3InputButton != null) {
                buildStage3InputButton.setTooltip(t);
            }
        } else {
            if (dispatchTrialButton != null) {
                dispatchTrialButton.setTooltip(null);
            }
            if (buildStage3InputButton != null) {
                buildStage3InputButton.setTooltip(
                        new Tooltip(
                                "保存済み 結果_配台表.json を枝番分解し『配台計画_タスク入力3.0』タブへ書き出す（段階3.0 の前処理）"));
            }
        }
        if (dispatchTrialButton != null) {
            if (deliveryCalendarReloadBlocking && !reloadInteractionDisabled) {
                dispatchTrialButton.setText(DISPATCH_TRIAL_BUTTON_TEXT_DELIVERY_CALENDAR_RELOAD);
            } else if (isSummaryExportLockedByLockFile() && !reloadInteractionDisabled) {
                dispatchTrialButton.setText(DISPATCH_TRIAL_BUTTON_TEXT_SUMMARY_LOCKED);
            } else {
                dispatchTrialButton.setText(DISPATCH_TRIAL_BUTTON_TEXT_DEFAULT);
            }
        }
        applyAlignToAladdinPlanButtonEnabledState();
    }

    private void applyAlignToAladdinPlanButtonEnabledState() {
        if (alignToAladdinPlanButton == null) {
            return;
        }
        boolean disabled =
                reloadInteractionDisabled
                        || deliveryCalendarReloadBlocking
                        || doc == null
                        || docHasActualDispatchQtyColumn()
                        || wideProfiles.isEmpty()
                        || dateAxis.isEmpty();
        alignToAladdinPlanButton.setDisable(disabled);
        if (deliveryCalendarReloadBlocking) {
            alignToAladdinPlanButton.setText(ALIGN_TO_ALADDIN_PLAN_BUTTON_TEXT_DELIVERY_CALENDAR_RELOAD);
            alignToAladdinPlanButton.setTooltip(
                    new Tooltip("納期管理ビューを再読み込み中です。完了後に実行してください。"));
        } else {
            alignToAladdinPlanButton.setText(ALIGN_TO_ALADDIN_PLAN_BUTTON_TEXT_DEFAULT);
            if (docHasActualDispatchQtyColumn()) {
                alignToAladdinPlanButton.setTooltip(
                        new Tooltip("段階3試行後は (段階3前) 数量の自動整列はできません"));
            } else {
                alignToAladdinPlanButton.setTooltip(
                        new Tooltip(
                                "段階3前: タスク×日付の (段階3前) 数量を (アラ計画) に沿うようロール単位で再配分する"));
            }
        }
    }

    /**
     * Wide-grid static columns (editable) push into {@link #doc}; column 0 (試行順) is not edited here (reordered via
     * DnD / buttons).
     */
    private void onWideGridChange(GridChange ev) {
        if (suppressDispatchGridDirty.get()) {
            return;
        }
        int r = ev.getRow();
        int c = ev.getColumn();
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        if (r < firstData) {
            return;
        }
        int staticCols = WIDE_STATIC_HEADERS.size();
        if (c <= 0 || c >= staticCols) {
            return;
        }
        if (isComputedWideStaticHeader(WIDE_STATIC_HEADERS.get(c))) {
            return;
        }
        int profileIdx = r - firstData;
        if (profileIdx < 0 || profileIdx >= wideProfiles.size()) {
            return;
        }
        Map<String, String> prof = wideProfiles.get(profileIdx);
        Map<String, String> oldProf = new LinkedHashMap<>(prof);
        String headerKey = WIDE_STATIC_HEADERS.get(c);
        Object nv = ev.getNewValue();
        String s = nv != null ? Objects.toString(nv, "") : "";
        prof.put(headerKey, s);
        List<String> cols = doc.columns();
        for (Map<String, String> row : doc.rows()) {
            if (ResultDispatchPivot.matchesTaskProfileExceptTrialOrder(cols, oldProf, row)) {
                row.put(headerKey, s);
            }
        }
        ResultDispatchNormalizer.normalizeInPlace(cols, doc.rows());
        markDispatchDocDirty();
    }

    private Path resolvePythonExeForShell(MainShellController shellRef) {
        if (shellRef == null) {
            return MainShellController.defaultPythonPathWhenShellMissing();
        }
        return shellRef.resolveStagePythonExecutablePath();
    }

    private Path resolvePythonExe() {
        return resolvePythonExeForShell(shell);
    }

    private void rebuildGrids() {
        rebuildGrids(null);
    }

    /**
     * 両 Spreadsheet の列同期・幅適用が終わったあと FX スレッドで {@code afterLayoutsReady} を実行する。
     * 同期前にモーダル（{@code showAndWait}）を出すと列数と View の不整合で {@link IndexOutOfBoundsException}
     * になることがある。
     */
    private void rebuildGrids(Runnable afterLayoutsReady) {
        if (dispatchReplay.isRunning()) {
            dispatchReplay.stop();
            setReplayButtonsRunning(false);
        }
                int layoutGen = bumpDispatchSpreadsheetLayoutGeneration();
        FullGridRebuild bundle = buildFullGridRebuild();
        applyFullGridRebuild(bundle, afterLayoutsReady, layoutGen);
    }

    private int bumpDispatchSpreadsheetLayoutGeneration() {
        return dispatchSpreadsheetLayoutGeneration.incrementAndGet();
    }

    private boolean isDispatchSpreadsheetLayoutStale(int layoutGen) {
        return layoutGen != dispatchSpreadsheetLayoutGeneration.get();
    }

    private List<LocalDate> computeDateAxisList() {
        if (aladdinPlanLookup.isEmpty()) {
            aladdinPlanLookup = loadAladdinPlanLookupForDisplay();
        }
        List<LocalDate> range =
                DispatchInteractiveDateAxis.computeInclusiveRange(
                        doc, aladdinPlanLookup, lastDispatchShortfallRows, dateAxisPastDays);
        if (range.isEmpty()) {
            return DispatchInteractiveDateAxis.defaultAxisWhenNoDataDates(dateAxisPastDays);
        }
        return range;
    }

    private FullGridRebuild buildFullGridRebuild() {
        ResultDispatchInteractiveGridModel.applyWideMergeAndNormalize(doc);
        aladdinPlanLookup = loadAladdinPlanLookupForDisplay();
        List<LocalDate> axis = axisForRebuild();
        WideGridBundle wide = buildWideGridModel(axis);
        ByDayGridBundle byDay = buildByDayGridModel(axis);
        return new FullGridRebuild(axis, wide, byDay);
    }

    private Map<String, Map<String, Map<String, Map<String, Double>>>> loadAladdinPlanLookupForDisplay() {
        Map<String, String> ui = shell != null ? shell.snapshotUiEnv() : Map.of();
        AladdinShapedPlanQtyLookup.ShapedTable shaped =
                AladdinShapedPlanQtyLookup.loadShapedTable(
                        AppPaths.resolveShapedAladdinPlanJsonPath(ui));
        return AladdinShapedPlanQtyLookup.buildLookup(shaped.headers(), shaped.rows());
    }

    /** データに存在する日付集合は維持しつつ、保存済み／ユーザー設定の列順を日付軸に反映する。 */
    private List<LocalDate> axisForRebuild() {
        List<LocalDate> computed = computeDateAxisList();
        if (preferredDateAxisOrder != null && sameMultisetLocalDate(preferredDateAxisOrder, computed)) {
            return new ArrayList<>(preferredDateAxisOrder);
        }
        preferredDateAxisOrder = null;
        List<LocalDate> fromPersistence = tryLoadPreferredDateOrderFromPersistence(computed);
        if (fromPersistence != null) {
            preferredDateAxisOrder = new ArrayList<>(fromPersistence);
            return new ArrayList<>(fromPersistence);
        }
        return computed;
    }

    private static boolean sameMultisetLocalDate(List<LocalDate> a, List<LocalDate> b) {
        if (a == null || b == null || a.size() != b.size()) {
            return false;
        }
        HashMap<LocalDate, Integer> freq = new HashMap<>();
        for (LocalDate d : a) {
            freq.merge(d, 1, Integer::sum);
        }
        for (LocalDate d : b) {
            Integer n = freq.get(d);
            if (n == null || n <= 0) {
                return false;
            }
            if (n == 1) {
                freq.remove(d);
            } else {
                freq.put(d, n - 1);
            }
        }
        return freq.isEmpty();
    }

    private List<LocalDate> tryLoadPreferredDateOrderFromPersistence(List<LocalDate> computed) {
        List<TableColumnOrderPersistence.ColumnSpec> lay =
                TableColumnOrderPersistence.loadLayout(TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_WIDE);
        if (lay == null || lay.isEmpty()) {
            return null;
        }
        List<String> titles = lay.stream().map(TableColumnOrderPersistence.ColumnSpec::title).toList();
        if (!wideStaticPrefixMatches(titles)
                || titles.size() != WIDE_STATIC_HEADERS.size() + computed.size()) {
            return null;
        }
        List<LocalDate> dates = parseDateTailAsDates(titles, WIDE_STATIC_HEADERS.size());
        if (dates == null || !sameMultisetLocalDate(dates, computed)) {
            return null;
        }
        return dates;
    }

    private static boolean wideStaticPrefixMatches(List<String> titles) {
        if (titles == null || titles.size() < WIDE_STATIC_HEADERS.size()) {
            return false;
        }
        for (int i = 0; i < WIDE_STATIC_HEADERS.size(); i++) {
            if (!WIDE_STATIC_HEADERS.get(i).equals(titles.get(i))) {
                return false;
            }
        }
        return true;
    }

    private boolean byDayStaticPrefixMatches(List<String> titles) {
        if (titles == null || titles.size() < BY_DAY_STATIC_HEADERS.size()) {
            return false;
        }
        for (int i = 0; i < BY_DAY_STATIC_HEADERS.size(); i++) {
            if (!BY_DAY_STATIC_HEADERS.get(i).equals(titles.get(i))) {
                return false;
            }
        }
        return true;
    }

    private static List<LocalDate> parseDateTailAsDates(List<String> titles, int staticCount) {
        if (titles == null || titles.size() < staticCount) {
            return null;
        }
        List<LocalDate> dates = new ArrayList<>();
        for (int i = staticCount; i < titles.size(); i++) {
            try {
                dates.add(LocalDate.parse(titles.get(i)));
            } catch (DateTimeParseException e) {
                return null;
            }
        }
        return dates;
    }

    private void persistDispatchColumnLayouts(List<String> wideTitles, List<String> byDayTitles) {
        List<TableColumnOrderPersistence.ColumnSpec> wideLay =
                TableColumnOrderPersistence.loadLayout(TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_WIDE);
        List<Double> wideW =
                TableColumnOrderPersistence.resolveWidthsForHeaders(wideTitles, wideLay, 112);
        List<TableColumnOrderPersistence.ColumnSpec> wideSpecs = new ArrayList<>();
        for (int i = 0; i < wideTitles.size(); i++) {
            wideSpecs.add(new TableColumnOrderPersistence.ColumnSpec(wideTitles.get(i), wideW.get(i)));
        }
        TableColumnOrderPersistence.saveLayout(TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_WIDE, wideSpecs);

        List<TableColumnOrderPersistence.ColumnSpec> byDayLay =
                TableColumnOrderPersistence.loadLayout(TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_BY_DAY);
        List<Double> byDayW =
                TableColumnOrderPersistence.resolveWidthsForHeaders(byDayTitles, byDayLay, 112);
        List<TableColumnOrderPersistence.ColumnSpec> byDaySpecs = new ArrayList<>();
        for (int i = 0; i < byDayTitles.size(); i++) {
            byDaySpecs.add(new TableColumnOrderPersistence.ColumnSpec(byDayTitles.get(i), byDayW.get(i)));
        }
        TableColumnOrderPersistence.saveLayout(
                TableColumnOrderPersistence.TableId.DISPATCH_INTERACTIVE_BY_DAY, byDaySpecs);
    }

    private void onWideSpreadsheetVisualColumnOrderChanged(List<String> titles) {
        if (suppressColumnReorderPersistence.get()) {
            return;
        }
        if (!wideStaticPrefixMatches(titles)) {
            return;
        }
        List<LocalDate> computed = computeDateAxisList();
        List<LocalDate> dates = parseDateTailAsDates(titles, WIDE_STATIC_HEADERS.size());
        if (dates == null || !sameMultisetLocalDate(dates, computed)) {
            return;
        }
        if (dates.equals(preferredDateAxisOrder)) {
            return;
        }
        preferredDateAxisOrder = new ArrayList<>(dates);
        persistDispatchColumnLayouts(titles, buildByDayColumnLabelsForAxis(dates));
        suppressColumnReorderPersistence.set(true);
        try {
            rebuildGrids();
        } finally {
            suppressColumnReorderPersistence.set(false);
        }
    }

    private void onByDaySpreadsheetVisualColumnOrderChanged(List<String> titles) {
        if (suppressColumnReorderPersistence.get()) {
            return;
        }
        if (!byDayStaticPrefixMatches(titles)) {
            return;
        }
        List<LocalDate> computed = computeDateAxisList();
        List<LocalDate> dates = parseDateTailAsDates(titles, BY_DAY_STATIC_HEADERS.size());
        if (dates == null || !sameMultisetLocalDate(dates, computed)) {
            return;
        }
        if (dates.equals(preferredDateAxisOrder)) {
            return;
        }
        preferredDateAxisOrder = new ArrayList<>(dates);
        persistDispatchColumnLayouts(buildWideColumnLabelsForAxis(dates), titles);
        suppressColumnReorderPersistence.set(true);
        try {
            rebuildGrids();
        } finally {
            suppressColumnReorderPersistence.set(false);
        }
    }

    private WideGridBundle buildWideGridModel(List<LocalDate> axis) {
        List<Map<String, String>> profiles = new ArrayList<>();
        List<WideRow> rowItems = new ArrayList<>();
        List<String> cols = doc.columns();
        Map<String, DispatchQtyShortfallRow> wideShortfallMerged =
                mergeWideDispatchShortfallRowsByKey();
        profiles.addAll(
                ResultDispatchPivot.distinctWideTaskProfiles(
                        cols,
                        doc.rows(),
                        ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS));
        profiles.sort(
                Comparator.comparing(DispatchInteractiveTabController::parseTrialOrderKey)
                        .thenComparing(p -> ResultDispatchNormalizer.staticGroupKey(cols, p)));
        assignSequentialTrialOrdersForProfiles(profiles);
        if (pendingStage3TrialSnapshotCapture) {
            captureStage3TrialPlanQtySnapshot(profiles, axis);
            pendingStage3TrialSnapshotCapture = false;
        }

        int staticCols = WIDE_STATIC_HEADERS.size();
        int dayCount = axis.size();
        int slotCols = dayCount * DAY_SLOT_COLUMNS;
        int totalCols = staticCols + slotCols;
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        int gridRowsTotal = firstData + profiles.size();
        GridBase grid = new GridBase(gridRowsTotal, totalCols);
        grid.getColumnHeaders().clear();
        grid.getColumnHeaders().addAll(buildWideColumnLabelsForAxis(axis));

        List<ObservableList<SpreadsheetCell>> gridRows = new ArrayList<>(gridRowsTotal);

        ObservableList<SpreadsheetCell> filterRow = FXCollections.observableArrayList();
        for (int c = 0; c < totalCols; c++) {
            SpreadsheetCell cell =
                    SpreadsheetCellType.STRING.createCell(
                            SpreadsheetTabularSupport.SPREADSHEET_FILTER_ROW, c, 1, 1, "");
            cell.setEditable(false);
            cell.setStyle(SpreadsheetTabularSupport.READABLE_STYLE_FILTER_ROW);
            filterRow.add(cell);
        }
        gridRows.add(filterRow);

        for (int pr = 0; pr < profiles.size(); pr++) {
            Map<String, String> profile = profiles.get(pr);
            int gridRow = firstData + pr;
            WideRow wr = new WideRow(profile, axis.size());
            for (int j = 0; j < axis.size(); j++) {
                LocalDate day = axis.get(j);
                wr.setAmount(
                        j,
                        wideDisplayPlanMetersForDate(profile, day));
                wr.setActualAmount(
                        j,
                        wideDisplayActualMetersForDate(profile, day));
            }
            rowItems.add(wr);

            ObservableList<SpreadsheetCell> line = FXCollections.observableArrayList();
            for (int c = 0; c < staticCols; c++) {
                String title = WIDE_STATIC_HEADERS.get(c);
                String raw = wideStaticCellText(wr, title);
                SpreadsheetCell cell =
                        SpreadsheetCellType.STRING.createCell(gridRow, c, 1, 1, raw != null ? raw : "");
                cell.setEditable(c > 0 && !isComputedWideStaticHeader(title));
                if (Stage3DispatchQtyBalanceCheck.COL_TITLE.equals(title)
                        && Stage3DispatchQtyBalanceCheck.isNgResult(raw)) {
                    cell.setStyle(TabularCellHighlight.PLAN_INPUT_EXCLUDE_YES_STYLE);
                } else {
                    cell.setStyle(SpreadsheetTabularSupport.READABLE_STYLE_LEADING_COL);
                }
                line.add(cell);
            }
            for (int di = 0; di < dayCount; di++) {
                int col = staticCols + di * DAY_SLOT_COLUMNS;
                SpreadsheetCell cell =
                        SpreadsheetCellType.STRING.createCell(gridRow, col, 1, 1, "");
                applyWideDispatchQtyCellDisplay(
                        cell, wr, di, axis, wideShortfallMerged);
                applyWideCellStyle(wr, di, cell);
                cell.setEditable(false);
                line.add(cell);
            }
            gridRows.add(line);
        }
        grid.setRows(gridRows);
        applyDispatchStage3QtyRowPresentation(grid);

        boolean[] wideBlockedCols = computeWideFullyBlockedDateColumns(dayCount);
        return new WideGridBundle(grid, profiles, rowItems, wideBlockedCols, staticCols, dayCount);
    }

    private ByDayGridBundle buildByDayGridModel(List<LocalDate> axis) {
        List<Map.Entry<String, String>> keys = ResultDispatchPivot.sortedProcessMachineKeys(doc.rows());
        int staticCols = BY_DAY_STATIC_HEADERS.size();
        int dayCount = axis.size();
        int slotCols = dayCount * DAY_SLOT_COLUMNS;
        int totalCols = staticCols + slotCols;
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        int gridRowsTotal = firstData + keys.size();
        GridBase grid = new GridBase(gridRowsTotal, totalCols);
        grid.getColumnHeaders().clear();
        grid.getColumnHeaders().addAll(buildByDayColumnLabelsForAxis(axis));

        List<ObservableList<SpreadsheetCell>> gridRows = new ArrayList<>(gridRowsTotal);

        ObservableList<SpreadsheetCell> filterRow = FXCollections.observableArrayList();
        for (int c = 0; c < totalCols; c++) {
            SpreadsheetCell cell =
                    SpreadsheetCellType.STRING.createCell(
                            SpreadsheetTabularSupport.SPREADSHEET_FILTER_ROW, c, 1, 1, "");
            cell.setEditable(false);
            cell.setStyle(SpreadsheetTabularSupport.READABLE_STYLE_FILTER_ROW);
            filterRow.add(cell);
        }
        gridRows.add(filterRow);

        List<String> cols = doc.columns();
        List<ByDayRow> byItems = new ArrayList<>();
        for (Map.Entry<String, String> en : keys) {
            String pcSummary =
                    ResultDispatchPivot.processingContentSummaryForProcessMachine(
                            cols, doc.rows(), en.getKey(), en.getValue());
            ByDayRow br = new ByDayRow(en.getKey(), en.getValue(), pcSummary, axis.size());
            for (int j = 0; j < axis.size(); j++) {
                LocalDate day = axis.get(j);
                br.setAmount(
                        j,
                        byDayDisplayPlanMetersForDate(en.getKey(), en.getValue(), day));
                br.setActualAmount(
                        j,
                        byDayDisplayActualMetersForDate(en.getKey(), en.getValue(), day));
            }
            byItems.add(br);
        }

        for (int ir = 0; ir < byItems.size(); ir++) {
            ByDayRow br = byItems.get(ir);
            int gridRow = firstData + ir;
            ObservableList<SpreadsheetCell> line = FXCollections.observableArrayList();
            for (int c = 0; c < staticCols; c++) {
                String title = BY_DAY_STATIC_HEADERS.get(c);
                String raw = byDayStaticCellText(br, title);
                SpreadsheetCell cell =
                        SpreadsheetCellType.STRING.createCell(gridRow, c, 1, 1, raw != null ? raw : "");
                cell.setEditable(false);
                cell.setStyle(SpreadsheetTabularSupport.READABLE_STYLE_LEADING_COL);
                line.add(cell);
            }
            for (int di = 0; di < dayCount; di++) {
                int col = staticCols + di * DAY_SLOT_COLUMNS;
                SpreadsheetCell cell =
                        SpreadsheetCellType.STRING.createCell(gridRow, col, 1, 1, "");
                applyByDayDispatchQtyCellDisplay(cell, br, di, axis);
                applyByDayCellStyle(br, di, cell);
                cell.setEditable(false);
                line.add(cell);
            }
            gridRows.add(line);
        }
        grid.setRows(gridRows);
        applyDispatchStage3QtyRowPresentation(grid);

        boolean[] byDayBlockedCols = computeByDayFullyBlockedDateColumns(dayCount);
        return new ByDayGridBundle(grid, byDayBlockedCols, staticCols, dayCount);
    }

    private void applyFullGridRebuild(
            FullGridRebuild bundle, Runnable afterLayoutsReady, int layoutGen) {
        suppressDispatchGridDirty.set(true);
        boolean emptyWide = bundle.wide().profiles().isEmpty();
        Map<Integer, Set<String>> wideColumnFilterSnapshot =
                emptyWide
                        ? Map.of()
                        : SpreadsheetMultiColumnFilterCoordinator.copyColumnAllowedByIndex(
                                wideSpreadsheet);
        Map<Integer, Set<String>> byDayColumnFilterSnapshot =
                emptyWide
                        ? Map.of()
                        : SpreadsheetMultiColumnFilterCoordinator.copyColumnAllowedByIndex(
                                byDaySpreadsheet);
        clearSpreadsheetSelectionForRebuild(wideSpreadsheet);
        clearSpreadsheetSelectionForRebuild(byDaySpreadsheet);
        if (dispatchShortfallTable != null) {
            dispatchShortfallTable.getSelectionModel().clearSelection();
        }
        try {
            dateAxis.clear();
            dateAxis.addAll(bundle.axis());
            wideProfiles.clear();
            wideProfiles.addAll(bundle.wide().profiles());
            wideRowItems.clear();
            wideRowItems.addAll(bundle.wide().rowItems());
            applyAlignToAladdinPlanButtonEnabledState();

            WideGridBundle w = bundle.wide();
            w.grid().addEventHandler(GridChange.GRID_CHANGE_EVENT, this::onWideGridChange);
                        SpreadsheetTabularSupport.detachAndSetGrid(wideSpreadsheet, w.grid());
            wideSpreadsheet.setFilteredRow(SpreadsheetTabularSupport.SPREADSHEET_FILTER_ROW);
            
            ByDayGridBundle b = bundle.byDay();
            SpreadsheetTabularSupport.detachAndSetGrid(byDaySpreadsheet, b.grid());
            byDaySpreadsheet.setFilteredRow(SpreadsheetTabularSupport.SPREADSHEET_FILTER_ROW);

            final Runnable layoutsReadyCallback = afterLayoutsReady;
            Platform.runLater(
                    () -> {
                        if (isDispatchSpreadsheetLayoutStale(layoutGen)) {
                            return;
                        }
                        AtomicInteger pendingLayouts = new AtomicInteger(2);
                        Runnable bothLayoutsReady =
                                () -> {
                                    if (isDispatchSpreadsheetLayoutStale(layoutGen)) {
                                                                                return;
                                    }
                                    if (pendingLayouts.decrementAndGet() != 0) {
                                        return;
                                    }
                                    clearSpreadsheetSelectionForRebuild(wideSpreadsheet);
                                    clearSpreadsheetSelectionForRebuild(byDaySpreadsheet);
                                    Platform.runLater(
                                            () -> {
                                                if (isDispatchSpreadsheetLayoutStale(layoutGen)) {
                                                    return;
                                                }
                                                                                                finalizeDispatchSpreadsheetPresentation();
                                                if (layoutsReadyCallback != null) {
                                                    layoutsReadyCallback.run();
                                                }
                                                Platform.runLater(
                                                        () -> {
                                                            if (wideSpreadsheetHost != null) {
                                                                wideSpreadsheetHost.requestLayout();
                                                            }
                                                            if (byDaySpreadsheetHost != null) {
                                                                byDaySpreadsheetHost
                                                                        .requestLayout();
                                                            }
                                                        });
                                            });
                                };
                        scheduleWideLayoutAfterColumnSync(
                                w, bothLayoutsReady, layoutGen, wideColumnFilterSnapshot);
                        scheduleByDayLayoutAfterColumnSync(
                                b, bothLayoutsReady, layoutGen, byDayColumnFilterSnapshot);
                    });

        } finally {
            Platform.runLater(() -> suppressDispatchGridDirty.set(false));
        }
    }

    /**
     * 段階2直後: タスク入力（配台不要オフ）の (依頼NO, 工程, 機械) が手動修正 JSON に無ければエラー表示。
     */
    private void showPlanInputCoverageGapErrorIfNeeded(Path resultDispatchJsonPath) {
        if (shell == null) {
            return;
        }
        PlanInputTabController planInput = shell.planInputTabControllerForDispatchRollUnit();
        if (planInput == null) {
            return;
        }
        List<TaskKey> expected = planInput.collectEligibleTaskKeysForDispatchCoverage();
        if (expected.isEmpty()) {
            return;
        }
        List<TaskKey> missing =
                DispatchPlanInputInteractiveCoverageCheck.findMissingInDispatchTable(
                        expected, doc != null ? doc.rows() : List.of());
        if (missing.isEmpty()) {
            return;
        }
        String pathStr = resultDispatchJsonPath != null ? resultDispatchJsonPath.toString() : "";
        String message =
                DispatchPlanInputInteractiveCoverageCheck.formatMissingTasksDialogMessage(
                        missing, pathStr);
        shell.appendLog(
                "[配台整合] タスク入力 "
                        + expected.size()
                        + " 件のうち手動修正表に無い行が "
                        + missing.size()
                        + " 件あります。");
        int logLines = 0;
        for (TaskKey key : missing) {
            if (logLines >= 15) {
                shell.appendLog("[配台整合] … 他 " + (missing.size() - 15) + " 件");
                break;
            }
            shell.appendLog("[配台整合] 未反映: " + key.displayLine());
            logLines++;
        }
        shell.showErrorDialog("手動修正表への反映漏れ", message);
    }

    private Runnable buildReloadSuccessAfterLayoutsRunnable(
            Runnable afterSuccessOnFxThread,
            boolean userCompletionDialog,
            boolean stage2ColsFilled,
            Path jsonPath) {
        if (afterSuccessOnFxThread == null && !userCompletionDialog) {
            return null;
        }
        return () -> {
            hideReloadProgress();
            if (afterSuccessOnFxThread != null) {
                afterSuccessOnFxThread.run();
            }
            if (userCompletionDialog) {
                String extra =
                        stage2ColsFilled
                                ? "\n\n段階2必須列を補完しました。必要に応じて保存してください。"
                                : "";
                shell.showInformationDialog(
                        "再読み完了",
                        doc.rows().size() + " 行を読み込みました。\n" + jsonPath + extra);
            }
                    };
    }

    private static void clearSpreadsheetSelectionForRebuild(SpreadsheetView view) {
        if (view == null) {
            return;
        }
        try {
            view.getSelectionModel().clearSelection();
        } catch (RuntimeException ignored) {
            // setGrid 直後など、選択モデルと列数が一時的にずれているときは無視する
        }
    }

    /** グリッド再構築後: 列フィルタは維持し、固定列・UNCONSTRAINED 列幅ポリシーを再適用する。 */
    private void finalizeDispatchSpreadsheetPresentation() {
        SpreadsheetTabularSupport.reapplySpreadsheetColumnChrome(
                wideSpreadsheet, resolvedWideLeadingColumnCount());
        SpreadsheetTabularSupport.reapplySpreadsheetColumnChrome(
                byDaySpreadsheet, resolvedByDayLeadingColumnCount());
    }

    private void refreshDispatchSpreadsheetForView(SpreadsheetView view) {
        if (view == null || view.getScene() == null) {
            return;
        }
                SpreadsheetTabularSupport.refreshSpreadsheetAfterRowPresentationChange(view, false);
    }


    
    /**
     * After {@link SpreadsheetView#setGrid}, the inner {@link TableView} may add columns on the next layout pulse.
     * Retrying avoids applying widths while {@link SpreadsheetView#getColumns()} is still shorter than the grid
     * (which skipped date columns and looked like “no date columns”).
     */
    private void scheduleWideLayoutAfterColumnSync(
            WideGridBundle w,
            Runnable onComplete,
            int layoutGen,
            Map<Integer, Set<String>> columnFilterSnapshot) {
        final int expectedCols = w.staticCols() + w.dayCount() * DAY_SLOT_COLUMNS;
        final int[] attempts = {0};
        final Runnable[] job = new Runnable[1];
        job[0] =
                () -> {
                    if (isDispatchSpreadsheetLayoutStale(layoutGen)) {
                                                return;
                    }
                    attempts[0]++;
                    int actual = wideSpreadsheet.getColumns().size();
                    boolean retry = actual < expectedCols && attempts[0] < 48;
                    if (retry) {
                        Platform.runLater(job[0]);
                        return;
                    }
                    Platform.runLater(
                            () -> {
                                if (isDispatchSpreadsheetLayoutStale(layoutGen)) {
                                    return;
                                }
                                wideSpreadsheet.setFilteredRow(
                                        SpreadsheetTabularSupport.SPREADSHEET_FILTER_ROW);
                                SpreadsheetTabularSupport.applyColumnFiltersWithDialog(
                                        wideSpreadsheet);
                                SpreadsheetMultiColumnFilterCoordinator.restoreColumnAllowedSnapshot(
                                        wideSpreadsheet, columnFilterSnapshot);
                                suppressDispatchColumnLayoutPersistence.set(true);
                                try {
                                    applyPersistedDispatchColumnWidths(
                                            wideSpreadsheet,
                                            buildWideColumnLabelsForAxis(dateAxis),
                                            TableColumnOrderPersistence.TableId
                                                    .DISPATCH_INTERACTIVE_WIDE);
                                    applyDateColumnWidthsForBlockedDays(
                                            wideSpreadsheet,
                                            w.staticCols(),
                                            w.dayCount(),
                                            sanitizeFullyBlockedFlagsForColumnWidth(
                                                    w.blockedCols()));
                                } finally {
                                    suppressDispatchColumnLayoutPersistence.set(false);
                                }
                                SpreadsheetTabularSupport.applyFixedLeadingColumns(
                                        wideSpreadsheet, resolvedWideLeadingColumnCount());
                                SpreadsheetTabularSupport.pinSpreadsheetFilterRow(wideSpreadsheet);
                                SpreadsheetTabularSupport
                                        .applyUnconstrainedColumnResizePolicyAfterSkinSettles(
                                                wideSpreadsheet);
                                SpreadsheetColumnDragReorderSupport.refreshAfterGridReady(
                                        wideSpreadsheet,
                                        suppressColumnReorderPersistence::get,
                                        () -> new ArrayList<>(buildWideColumnLabelsForAxis(dateAxis)),
                                        WIDE_STATIC_HEADERS.size(),
                                        this::onWideSpreadsheetVisualColumnOrderChanged);
                                ColumnVisibilitySupport.applyColumnVisibilityToSpreadsheetWhenReady(
                                        wideSpreadsheet,
                                        () -> new ArrayList<>(buildWideColumnLabelsForAxis(dateAxis)),
                                        () -> {
                                            List<String> h = buildWideColumnLabelsForAxis(dateAxis);
                                            return TableColumnOrderPersistence.loadColumnVisibility(
                                                    TableColumnOrderPersistence.TableId
                                                            .DISPATCH_INTERACTIVE_WIDE,
                                                    h.size());
                                        },
                                        () ->
                                                runAfterWideSpreadsheetLayoutSettled(
                                                        onComplete, layoutGen));
                            });
                };
        Platform.runLater(job[0]);
    }

    private void runAfterWideSpreadsheetLayoutSettled(Runnable onComplete, int layoutGen) {
        Platform.runLater(
                () -> {
                    if (isDispatchSpreadsheetLayoutStale(layoutGen)) {
                        return;
                    }
                    SpreadsheetTabularSupport.reapplySpreadsheetColumnChrome(
                            wideSpreadsheet, resolvedWideLeadingColumnCount());
                    if (onComplete != null) {
                        onComplete.run();
                    } else {
                        refreshDispatchSpreadsheetForView(wideSpreadsheet);
                    }
                });
    }

    private void scheduleByDayLayoutAfterColumnSync(
            ByDayGridBundle b,
            Runnable onComplete,
            int layoutGen,
            Map<Integer, Set<String>> columnFilterSnapshot) {
        final int expectedCols = b.staticCols() + b.dayCount() * DAY_SLOT_COLUMNS;
        final int[] attempts = {0};
        final Runnable[] job = new Runnable[1];
        job[0] =
                () -> {
                    if (isDispatchSpreadsheetLayoutStale(layoutGen)) {
                                                return;
                    }
                    attempts[0]++;
                    if (byDaySpreadsheet.getColumns().size() < expectedCols && attempts[0] < 48) {
                        Platform.runLater(job[0]);
                        return;
                    }
                    Platform.runLater(
                            () -> {
                                if (isDispatchSpreadsheetLayoutStale(layoutGen)) {
                                    return;
                                }
                                byDaySpreadsheet.setFilteredRow(
                                        SpreadsheetTabularSupport.SPREADSHEET_FILTER_ROW);
                                SpreadsheetTabularSupport.applyColumnFiltersWithDialog(
                                        byDaySpreadsheet);
                                SpreadsheetMultiColumnFilterCoordinator.restoreColumnAllowedSnapshot(
                                        byDaySpreadsheet, columnFilterSnapshot);
                                suppressDispatchColumnLayoutPersistence.set(true);
                                try {
                                    applyPersistedDispatchColumnWidths(
                                            byDaySpreadsheet,
                                            buildByDayColumnLabelsForAxis(dateAxis),
                                            TableColumnOrderPersistence.TableId
                                                    .DISPATCH_INTERACTIVE_BY_DAY);
                                    applyDateColumnWidthsForBlockedDays(
                                            byDaySpreadsheet,
                                            b.staticCols(),
                                            b.dayCount(),
                                            sanitizeFullyBlockedFlagsForColumnWidth(
                                                    b.blockedCols()));
                                } finally {
                                    suppressDispatchColumnLayoutPersistence.set(false);
                                }
                                SpreadsheetTabularSupport.applyFixedLeadingColumns(
                                        byDaySpreadsheet, resolvedByDayLeadingColumnCount());
                                SpreadsheetTabularSupport.pinSpreadsheetFilterRow(byDaySpreadsheet);
                                SpreadsheetTabularSupport
                                        .applyUnconstrainedColumnResizePolicyAfterSkinSettles(
                                                byDaySpreadsheet);
                                SpreadsheetColumnDragReorderSupport.refreshAfterGridReady(
                                        byDaySpreadsheet,
                                        suppressColumnReorderPersistence::get,
                                        () -> new ArrayList<>(buildByDayColumnLabelsForAxis(dateAxis)),
                                        BY_DAY_STATIC_HEADERS.size(),
                                        this::onByDaySpreadsheetVisualColumnOrderChanged);
                                ColumnVisibilitySupport.applyColumnVisibilityToSpreadsheetWhenReady(
                                        byDaySpreadsheet,
                                        () -> new ArrayList<>(buildByDayColumnLabelsForAxis(dateAxis)),
                                        () -> {
                                            List<String> h = buildByDayColumnLabelsForAxis(dateAxis);
                                            return TableColumnOrderPersistence.loadColumnVisibility(
                                                    TableColumnOrderPersistence.TableId
                                                            .DISPATCH_INTERACTIVE_BY_DAY,
                                                    h.size());
                                        },
                                        () ->
                                                runAfterByDaySpreadsheetLayoutSettled(
                                                        onComplete, layoutGen));
                            });
                };
        Platform.runLater(job[0]);
    }

    private void runAfterByDaySpreadsheetLayoutSettled(Runnable onComplete, int layoutGen) {
        Platform.runLater(
                () -> {
                    if (isDispatchSpreadsheetLayoutStale(layoutGen)) {
                        return;
                    }
                    SpreadsheetTabularSupport.reapplySpreadsheetColumnChrome(
                            byDaySpreadsheet, resolvedByDayLeadingColumnCount());
                    if (onComplete != null) {
                        onComplete.run();
                    } else {
                        refreshDispatchSpreadsheetForView(byDaySpreadsheet);
                    }
                });
    }

    /**
     * When every date column is “fully blocked”, narrowing all of them to ~5pt makes the timeline disappear. Keep
     * default widths in that case (still gray cells via styles).
     */
    private static boolean[] sanitizeFullyBlockedFlagsForColumnWidth(boolean[] fullyBlocked) {
        if (fullyBlocked == null || fullyBlocked.length == 0) {
            return fullyBlocked;
        }
        int trueCount = 0;
        for (boolean b : fullyBlocked) {
            if (b) {
                trueCount++;
            }
        }
        if (trueCount == fullyBlocked.length) {
            return new boolean[fullyBlocked.length];
        }
        return fullyBlocked;
    }

    private List<String> buildWideColumnLabelsForAxis(List<LocalDate> axis) {
        List<String> headers =
                new ArrayList<>(WIDE_STATIC_HEADERS.size() + axis.size() * DAY_SLOT_COLUMNS);
        headers.addAll(WIDE_STATIC_HEADERS);
        for (LocalDate d : axis) {
            headers.add(d.toString());
        }
        return headers;
    }

    private List<String> buildByDayColumnLabelsForAxis(List<LocalDate> axis) {
        List<String> headers =
                new ArrayList<>(BY_DAY_STATIC_HEADERS.size() + axis.size() * DAY_SLOT_COLUMNS);
        headers.addAll(BY_DAY_STATIC_HEADERS);
        for (LocalDate d : axis) {
            headers.add(d.toString());
        }
        return headers;
    }

    private static double pointsToLocalPixels(double pt) {
        Screen s = Screen.getPrimary();
        if (s == null) {
            return pt * 96.0 / 72.0;
        }
        return pt * s.getDpi() / 72.0;
    }

    /**
     * Date columns where every profile row is blocked on that day get a narrow width; mixed columns stay default.
     */
    private boolean[] computeWideFullyBlockedDateColumns(int dayCount) {
        return new boolean[dayCount];
    }

    private boolean[] computeByDayFullyBlockedDateColumns(int dayCount) {
        return new boolean[dayCount];
    }

    private static void applyDateColumnWidthsForBlockedDays(
            SpreadsheetView view, int staticCols, int dayCount, boolean[] fullyBlocked) {
        if (view == null || fullyBlocked == null || dayCount <= 0) {
            return;
        }
        var cols = view.getColumns();
        if (cols.isEmpty()) {
            return;
        }
        double narrowPt = pointsToLocalPixels(BLOCKED_DATE_COLUMN_PREF_PT);
        double narrow = Math.max(narrowPt, MIN_BLOCKED_DATE_COLUMN_WIDTH_PX);
        for (int di = 0; di < dayCount; di++) {
            int colIndex = staticCols + di * DAY_SLOT_COLUMNS;
            if (colIndex >= cols.size()) {
                continue;
            }
            SpreadsheetColumn sc = cols.get(colIndex);
            if (fullyBlocked[di]) {
                sc.setPrefWidth(narrow);
                sc.setMinWidth(narrow);
                sc.setMaxWidth(narrow);
            } else {
                sc.setMinWidth(MIN_DATE_COLUMN_WIDTH_PX);
                sc.setMaxWidth(Double.MAX_VALUE);
            }
        }
    }

    private void applyWideCellStyle(WideRow wr, int dateIdx, SpreadsheetCell cell) {
        if (isWideDispatchShortfall(wr, dateIdx)) {
            cell.setStyle(DATE_CELL_STYLE_SHORTFALL);
            return;
        }
        String alignStyle = aladdinAlignMoveCellStyle(aladdinAlignMoveHighlightForWideRow(wr, dateIdx));
        if (alignStyle != null) {
            cell.setStyle(alignStyle);
            return;
        }
        double planQ = wr.getAmount(dateIdx);
        double actualQ = wr.getActualAmount(dateIdx);
        if (planQ > 1e-9 || actualQ > 1e-9) {
            cell.setStyle(DATE_CELL_STYLE_POSITIVE_QTY);
        } else {
            cell.setStyle(SpreadsheetTabularSupport.READABLE_STYLE_DATA_WHITE);
        }
    }

    private static double parseTrialOrderKey(Map<String, String> profile) {
        String s = profile.get(ResultDispatchSchema.COL_DISPATCH_TRIAL_ORDER);
        if (s == null || s.isBlank()) {
            return Double.MAX_VALUE;
        }
        try {
            return Double.parseDouble(s.trim().replace(",", ""));
        } catch (NumberFormatException e) {
            return Double.MAX_VALUE;
        }
    }

    /** Ensures each profile row maps to sequential trial order and pushes into doc rows. */
    private void assignSequentialTrialOrdersForProfiles(List<Map<String, String>> profiles) {
        for (int i = 0; i < profiles.size(); i++) {
            String ord = Integer.toString(i + 1);
            Map<String, String> prof = profiles.get(i);
            prof.put(ResultDispatchSchema.COL_DISPATCH_TRIAL_ORDER, ord);
            for (Map<String, String> row : doc.rows()) {
                if (ResultDispatchPivot.matchesWideMergeIdentity(
                        prof,
                        row,
                        ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS)) {
                    row.put(ResultDispatchSchema.COL_DISPATCH_TRIAL_ORDER, ord);
                }
            }
        }
    }

    private void swapWideProfiles(int a, int b) {
        Map<String, String> pa = wideProfiles.get(a);
        Map<String, String> pb = wideProfiles.get(b);
        wideProfiles.set(a, pb);
        wideProfiles.set(b, pa);
        assignSequentialTrialOrdersForProfiles(wideProfiles);
        ResultDispatchNormalizer.normalizeInPlace(doc.columns(), doc.rows());
        rebuildGrids();
        markDispatchDocDirty();
    }

    /** Column index in {@link SpreadsheetView#getColumns()} for the focused / primary selected cell. */
    private int wideSpreadsheetFocusedColumnIndex() {
        var sm = wideSpreadsheet.getSelectionModel();
        TablePosition<?, ?> pos = sm.getFocusedCell();
        if (pos != null && pos.getColumn() >= 0) {
            return pos.getColumn();
        }
        var cells = sm.getSelectedCells();
        if (cells != null && !cells.isEmpty()) {
            int c = cells.getFirst().getColumn();
            if (c >= 0) {
                return c;
            }
        }
        return 0;
    }

    /**
     * After row reorder, move selection and focus to the same logical profile row (by index) and column.
     * Runs later so spreadsheet layout is applied before selecting.
     */
    private void focusWideProfileCellAfterReorder(int profileIndex, int columnIndex) {
        if (profileIndex < 0 || profileIndex >= wideProfiles.size()) {
            return;
        }
        var cols = wideSpreadsheet.getColumns();
        if (cols.isEmpty()) {
            return;
        }
        int c = Math.max(0, Math.min(columnIndex, cols.size() - 1));
        SpreadsheetColumn scol = cols.get(c);
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        int modelGridRow = firstData + profileIndex;
        Platform.runLater(
                () -> {
                    int viewRow = wideSpreadsheet.getViewRow(modelGridRow);
                    if (viewRow < 0) {
                        return;
                    }
                    var sm = wideSpreadsheet.getSelectionModel();
                    sm.clearSelection();
                    sm.clearAndSelect(viewRow, scol);
                    sm.focus(viewRow, scol);
                    scrollWideSpreadsheetCellIntoView(viewRow, scol);
                });
    }

    /**
     * 未達サマリ表で選択した依頼NO・機械・配台日に対応するワイドグリッドのセルを選択・フォーカスする。
     */
    /** 「配台リプレイ」ボタン: 表示中の配台結果を加工開始日時順にワイド表で順次ハイライト再生する。 */
    @FXML
    private void onDispatchReplayAction() {
        if (dispatchReplay.isRunning()) {
            return;
        }
        if (doc == null || doc.rows().isEmpty() || wideProfiles.isEmpty() || dateAxis.isEmpty()) {
            setStatusSafe("配台リプレイ: 表示中の配台結果がありません。先に配台結果を読み込んでください。");
            return;
        }
        List<jp.co.pm.ai.desktop.dispatch.DispatchReplayController.Step> steps =
                jp.co.pm.ai.desktop.dispatch.DispatchReplayController.buildStepsFromRows(doc.rows());
        if (steps.isEmpty()) {
            setStatusSafe("配台リプレイ: 再生対象（当日配台数量 > 0）がありません。");
            return;
        }
        applyInnerTabSelectedIndex(0);
        if (dispatchReplaySpeedSlider != null) {
            dispatchReplay.setIntervalMillis(dispatchReplaySpeedSlider.getValue());
        }
        dispatchReplay.load(steps);
        setReplayButtonsRunning(true);
        dispatchReplay.play(
                this::visitDispatchReplayStep,
                () -> {
                    setReplayButtonsRunning(false);
                    setStatusSafe("配台リプレイ完了（" + steps.size() + " ステップ）。");
                });
    }

    /** 「停止」ボタン: 配台リプレイを中止する。 */
    @FXML
    private void onDispatchReplayStopAction() {
        dispatchReplay.stop();
        setReplayButtonsRunning(false);
        setStatusSafe("配台リプレイを停止しました。");
    }

    private void setReplayButtonsRunning(boolean running) {
        if (dispatchReplayButton != null) {
            dispatchReplayButton.setDisable(running);
        }
        if (dispatchReplayStopButton != null) {
            dispatchReplayStopButton.setDisable(!running);
        }
    }

    private void setStatusSafe(String msg) {
        if (statusLabel != null) {
            statusLabel.setText(msg);
        }
    }

    /** 1 ステップ再生: 対応するワイド表セルへフォーカス・スクロールし、ステータスへ進捗を出す。 */
    private void visitDispatchReplayStep(
            jp.co.pm.ai.desktop.dispatch.DispatchReplayController.Step step, int index, int total) {
        int profileIdx = findWideProfileIndexForRequestMachine(step.requestNo(), step.machine());
        int dateIdx = dateAxis.indexOf(step.dispatchDate());
        if (profileIdx >= 0 && dateIdx >= 0) {
            int modelCol = WIDE_STATIC_HEADERS.size() + dateIdx * DAY_SLOT_COLUMNS;
            focusWideProfileCellAfterReorder(profileIdx, modelCol);
        }
        setStatusSafe("配台リプレイ " + (index + 1) + "/" + total + ": " + step.label());
    }

    /** {@link #wideProfiles} を (依頼NO, 機械名) で照合する。見つからなければ -1。 */
    private int findWideProfileIndexForRequestMachine(String requestNo, String machine) {
        if (wideProfiles.isEmpty()) {
            return -1;
        }
        String req = requestNo == null ? "" : requestNo.trim();
        String mac = machine == null ? "" : machine.trim();
        for (int i = 0; i < wideProfiles.size(); i++) {
            Map<String, String> p = wideProfiles.get(i);
            String pReq = p.get("依頼NO"); // Literal Japanese key
            String pMac = p.get(ResultDispatchSchema.COL_MACHINE);
            if (req.equals(pReq == null ? "" : pReq.trim())
                    && mac.equals(pMac == null ? "" : pMac.trim())) {
                return i;
            }
        }
        return -1;
    }

    private void wireDispatchShortfallSelectionToWideGrid() {
        if (dispatchShortfallTable == null) {
            return;
        }
        dispatchShortfallTable
                .getSelectionModel()
                .selectedItemProperty()
                .addListener(
                        (obs, prev, row) -> {
                            focusWideSpreadsheetOnDispatchShortfallRow(row);
                        });
    }

    /**
     * {@link DispatchTrialShortages#wideShortfallKey} と同一規則でプロファイル行を特定する。
     *
     * @return {@link #wideProfiles} インデックス、無ければ -1
     */
    private int findWideProfileIndexMatchingShortfall(DispatchQtyShortfallRow row) {
        if (row == null || wideProfiles.isEmpty()) {
            return -1;
        }
        String expected =
                DispatchTrialShortages.wideShortfallKey(
                        row.taskId(), row.machineName(), row.dispatchDateIso());
        for (int i = 0; i < wideProfiles.size(); i++) {
            Map<String, String> p = wideProfiles.get(i);
            String candidate =
                    DispatchTrialShortages.wideShortfallKey(
                            p.get("依頼NO"),
                            p.get(ResultDispatchSchema.COL_MACHINE),
                            row.dispatchDateIso());
            if (expected.equals(candidate)) {
                return i;
            }
        }
        return -1;
    }

    /**
     * 列同期が直後に終わっていないときは数パルス待ってからフォーカスする（{@link #scheduleWideLayoutAfterColumnSync}
     * と同趣旨）。
     */
    private void scheduleFocusWideCellWhenShortfallReady(int profileIdx, int modelCol) {
        final int[] attempts = {0};
        final Runnable[] job = new Runnable[1];
        job[0] =
                () -> {
                    attempts[0]++;
                    var cols = wideSpreadsheet.getColumns();
                    boolean colsReady =
                            !cols.isEmpty()
                                    && modelCol >= 0
                                    && modelCol < cols.size()
                                    && profileIdx >= 0
                                    && profileIdx < wideProfiles.size();
                    if (!colsReady) {
                        if (attempts[0] < 48) {
                            Platform.runLater(job[0]);
                        }
                        return;
                    }
                    focusWideProfileCellAfterReorder(profileIdx, modelCol);
                };
        Platform.runLater(job[0]);
    }

    private void focusWideSpreadsheetOnDispatchShortfallRow(DispatchQtyShortfallRow row) {
        if (row == null || wideProfiles.isEmpty() || dateAxis.isEmpty()) {
            return;
        }
        int profileIdx = findWideProfileIndexMatchingShortfall(row);
        if (profileIdx < 0) {
            return;
        }
        LocalDate targetDate;
        try {
            targetDate = LocalDate.parse(row.dispatchDateIso().trim());
        } catch (DateTimeParseException e) {
            return;
        }
        int dateIdx = dateAxis.indexOf(targetDate);
        if (dateIdx < 0) {
            return;
        }
        int staticCols = WIDE_STATIC_HEADERS.size();
        int modelCol = staticCols + dateIdx * DAY_SLOT_COLUMNS;
        applyInnerTabSelectedIndex(0);
        scheduleFocusWideCellWhenShortfallReady(profileIdx, modelCol);
    }

    /** ControlsFX の {@link SpreadsheetView#scrollToRow} / {@link SpreadsheetView#scrollToColumn} で見える位置へ寄せる。 */
    private void scrollWideSpreadsheetCellIntoView(int viewRow, SpreadsheetColumn scol) {
        if (scol == null || viewRow < 0) {
            return;
        }
        wideSpreadsheet.scrollToRow(viewRow);
        wideSpreadsheet.scrollToColumn(scol);
    }

    private int selectedWideProfileIndex() {
        var sm = wideSpreadsheet.getSelectionModel();
        TablePosition<?, ?> pos = sm.getFocusedCell();
        if (pos == null || pos.getRow() < 0) {
            var cells = sm.getSelectedCells();
            if (cells == null || cells.isEmpty()) {
                return -1;
            }
            pos = cells.getFirst();
        }
        int viewRow = pos.getRow();
        int gridRow = wideSpreadsheet.getModelRow(viewRow);
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        int idx = gridRow - firstData;
        if (idx >= 0 && idx < wideProfiles.size()) {
            return idx;
        }
        return -1;
    }

    /** Maps a SpreadsheetView table/view row index to a {@link #wideProfiles} index. */
    private int wideProfileIndexFromViewRow(int viewRow) {
        if (viewRow < 0) {
            return -1;
        }
        int gridRow = wideSpreadsheet.getModelRow(viewRow);
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        int idx = gridRow - firstData;
        if (idx >= 0 && idx < wideProfiles.size()) {
            return idx;
        }
        return -1;
    }

    private int wideProfileIndexFromTableCell(TableCell<?, ?> tc) {
        if (tc == null) {
            return -1;
        }
        return wideProfileIndexFromViewRow(tc.getIndex());
    }

    /**
     * Maps a {@link TableCell}'s column to model column index (accounts for hidden columns / ControlsFX mapping).
     */
    private int wideModelColumnFromTableCell(TableCell<?, ?> tc) {
        if (tc == null || tc.getTableColumn() == null) {
            return -1;
        }
        int viewCol = tc.getTableView().getColumns().indexOf(tc.getTableColumn());
        if (viewCol < 0) {
            return -1;
        }
        return wideSpreadsheet.getModelColumn(viewCol);
    }

    private void applyByDayCellStyle(ByDayRow br, int dateIdx, SpreadsheetCell cell) {
        if (isByDayDispatchShortfall(br, dateIdx)) {
            cell.setStyle(DATE_CELL_STYLE_SHORTFALL);
            return;
        }
        String alignStyle = aladdinAlignMoveCellStyle(aladdinAlignMoveHighlightForByDay(br, dateIdx));
        if (alignStyle != null) {
            cell.setStyle(alignStyle);
            return;
        }
        double planQ = br.getAmount(dateIdx);
        double actualQ = br.getActualAmount(dateIdx);
        if (planQ > 1e-9 || actualQ > 1e-9) {
            cell.setStyle(DATE_CELL_STYLE_POSITIVE_QTY);
        } else {
            cell.setStyle(SpreadsheetTabularSupport.READABLE_STYLE_DATA_WHITE);
        }
    }

    private boolean isWideDispatchShortfall(WideRow wr, int dateIdx) {
        if (dispatchWideShortfallKeys.isEmpty()
                || dateIdx < 0
                || dateIdx >= dateAxis.size()) {
            return false;
        }
        String tid = wr.getStatic("依頼NO");
        String mach = wr.getStatic(ResultDispatchSchema.COL_MACHINE);
        LocalDate d = dateAxis.get(dateIdx);
        String key =
                DispatchTrialShortages.wideShortfallKey(tid, mach, d.toString());
        return dispatchWideShortfallKeys.contains(key);
    }

    private boolean isByDayDispatchShortfall(ByDayRow br, int dateIdx) {
        if (dispatchByDayShortfallKeys.isEmpty()
                || dateIdx < 0
                || dateIdx >= dateAxis.size()) {
            return false;
        }
        String mach = br.machine();
        LocalDate d = dateAxis.get(dateIdx);
        return dispatchByDayShortfallKeys.contains(
                DispatchTrialShortages.byDayShortfallKey(mach, d.toString()));
    }

    /**
     * {@code dispatch_qty_shortfall} をワイドキー（依頼NO・機械・配台日）で集約する。同一キーが複数あるときは目標・実績・不足を
     * 合算する。
     */
    private Map<String, DispatchQtyShortfallRow> mergeWideDispatchShortfallRowsByKey() {
        Map<String, double[]> sumByKey = new HashMap<>();
        Map<String, String[]> metaByKey = new HashMap<>();
        for (DispatchQtyShortfallRow r : lastDispatchShortfallRows) {
            String k =
                    DispatchTrialShortages.wideShortfallKey(
                            r.taskId(), r.machineName(), r.dispatchDateIso());
            sumByKey.merge(
                    k,
                    new double[] {r.targetM(), r.doneM(), r.shortfallM()},
                    (a, b) ->
                            new double[] {
                                a[0] + b[0], a[1] + b[1], a[2] + b[2]
                            });
            metaByKey.putIfAbsent(
                    k,
                    new String[] {
                        r.taskId() != null ? r.taskId() : "",
                        r.machineName() != null ? r.machineName() : "",
                        r.dispatchDateIso() != null ? r.dispatchDateIso() : ""
                    });
        }
        Map<String, DispatchQtyShortfallRow> out = new LinkedHashMap<>();
        for (Map.Entry<String, double[]> e : sumByKey.entrySet()) {
            String[] meta = metaByKey.get(e.getKey());
            double[] v = e.getValue();
            out.put(
                    e.getKey(),
                    new DispatchQtyShortfallRow(
                            meta[0], meta[1], meta[2], v[0], v[1], v[2], ""));
        }
        return out;
    }

    private boolean docHasActualDispatchQtyColumn() {
        return doc.columns().contains(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL);
    }

    private static boolean documentHasPositiveActualDispatchQty(ResultDispatchDocument document) {
        if (document == null || document.rows() == null) {
            return false;
        }
        if (!document.columns().contains(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL)) {
            return false;
        }
        for (Map<String, String> row : document.rows()) {
            if (row == null) {
                continue;
            }
            if (ResultDispatchNormalizer.parseDouble(
                            row.get(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL))
                    > 1e-6) {
                return true;
            }
        }
        return false;
    }

    /** 環境変数 {@link AppPaths#KEY_PM_AI_DEBUG_STAGE3_PLAN_ACTUAL_SINGLE_LINE}（未設定時は2行表示）。 */
    private boolean stage3PlanActualSingleLineDisplay() {
        if (shell == null) {
            return false;
        }
        return AppPaths.isTruthyUiEnv(
                shell.snapshotUiEnv(),
                AppPaths.KEY_PM_AI_DEBUG_STAGE3_PLAN_ACTUAL_SINGLE_LINE,
                false);
    }

    private static boolean isComputedWideStaticHeader(String title) {
        return COL_STAGE3_DISPATCH_QTY_TOTAL.equals(title)
                || Stage3DispatchQtyBalanceCheck.COL_TITLE.equals(title)
                || COL_SPECIAL_RULES.equals(title);
    }

    private String wideStaticCellText(WideRow wr, String title) {
        if (COL_SPECIAL_RULES.equals(title)) {
            return specialRuleBadges.badgeForRow(
                    wr.getStatic("依頼NO"),
                    wr.getStatic(ResultDispatchSchema.COL_PROCESS),
                    wr.getStatic(ResultDispatchSchema.COL_MACHINE));
        }
        if (COL_STAGE3_DISPATCH_QTY_TOTAL.equals(title)) {
            return formatStage3DispatchQtyTotal(stage3DispatchQtyTotalForWideRow(wr));
        }
        if (Stage3DispatchQtyBalanceCheck.COL_TITLE.equals(title)) {
            return formatStage3DispatchQtyBalanceCheck(wr);
        }
        return wr.getStatic(title);
    }

    private String formatStage3DispatchQtyBalanceCheck(WideRow wr) {
        double qtyConv = ResultDispatchNormalizer.parseDouble(wr.getStatic("換算数量"));
        double actualDone = ResultDispatchNormalizer.parseDouble(wr.getStatic("実加工数"));
        double rollUnitM = resolveRollUnitForWideRow(wr).unitM();
        double actualTotal = stage3DispatchQtyTotalForWideRow(wr);
        return Stage3DispatchQtyBalanceCheck.formatCheck(
                        qtyConv,
                        actualDone,
                        actualTotal,
                        docHasActualDispatchQtyColumn(),
                        rollUnitM);
    }

    /**
     * 段階3配台数・照合の合計。段階3試行後（実配台数量列あり）はタイムライン実績（段階3後）の日別合計。
     * 未試行時は当日配台数量の合計にフォールバック。
     */
    private double stage3DispatchQtyTotalForWideRow(WideRow wr) {
        if (docHasActualDispatchQtyColumn()) {
            return wr.sumActualAmounts();
        }
        return wr.sumPlanAmounts();
    }

    /**
     * 段階3試行後は設備ガント契約の暦日別 m を優先（配台表1行集約とガントの差を解消）。未読込時は配台表。
     */
    private double wideDisplayPlanMetersForDate(Map<String, String> profile, LocalDate day) {
        if (docHasActualDispatchQtyColumn() && timelineCalendarMeters.isLoaded()) {
            Optional<Double> tl =
                    timelineCalendarMeters.metersForTaskProfile(
                            profile.get("依頼NO"),
                            profile.get(ResultDispatchSchema.COL_PROCESS),
                            profile.get(ResultDispatchSchema.COL_MACHINE),
                            day);
            if (tl.isPresent()) {
                return tl.get();
            }
        }
        return ResultDispatchPivot.sumQuantityForProfileAndDateForWideMerge(
                doc.rows(),
                profile,
                day,
                ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
    }

    private double wideDisplayActualMetersForDate(Map<String, String> profile, LocalDate day) {
        if (docHasActualDispatchQtyColumn() && timelineCalendarMeters.isLoaded()) {
            Optional<Double> tl =
                    timelineCalendarMeters.metersForTaskProfile(
                            profile.get("依頼NO"),
                            profile.get(ResultDispatchSchema.COL_PROCESS),
                            profile.get(ResultDispatchSchema.COL_MACHINE),
                            day);
            if (tl.isPresent()) {
                return tl.get();
            }
        }
        return ResultDispatchPivot.sumActualQuantityForProfileAndDateForWideMerge(
                doc.rows(),
                profile,
                day,
                ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
    }

    private double byDayDisplayPlanMetersForDate(String process, String machine, LocalDate day) {
        if (docHasActualDispatchQtyColumn() && timelineCalendarMeters.isLoaded()) {
            Optional<Double> tl =
                    timelineCalendarMeters.metersForProcessMachine(process, machine, day);
            if (tl.isPresent()) {
                return tl.get();
            }
        }
        return ResultDispatchPivot.sumQuantityForProcessMachineDate(
                doc.rows(), process, machine, day);
    }

    private double byDayDisplayActualMetersForDate(String process, String machine, LocalDate day) {
        if (docHasActualDispatchQtyColumn() && timelineCalendarMeters.isLoaded()) {
            Optional<Double> tl =
                    timelineCalendarMeters.metersForProcessMachine(process, machine, day);
            if (tl.isPresent()) {
                return tl.get();
            }
        }
        return ResultDispatchPivot.sumActualQuantityForProcessMachineDate(
                doc.rows(), process, machine, day);
    }

    private double stage3DispatchQtyTotalForByDayRow(ByDayRow br) {
        if (docHasActualDispatchQtyColumn()) {
            return br.sumActualAmounts();
        }
        return br.sumPlanAmounts();
    }

    private String byDayStaticCellText(ByDayRow br, String title) {
        if (COL_STAGE3_DISPATCH_QTY_TOTAL.equals(title)) {
            return formatStage3DispatchQtyTotal(stage3DispatchQtyTotalForByDayRow(br));
        }
        if (ResultDispatchSchema.COL_PROCESS.equals(title)) {
            return br.process();
        }
        if (ResultDispatchSchema.COL_MACHINE.equals(title)) {
            return br.machine();
        }
        if ("加工内容".equals(title)) {
            return br.processingContent();
        }
        return "";
    }

    private String formatStage3DispatchQtyTotal(double sumQty) {
        if ((!docHasActualDispatchQtyColumn() && !hasPipelineStage3PlanningApplied())
                || sumQty <= 1e-3) {
            return "";
        }
        return ResultDispatchNormalizer.formatQty(sumQty);
    }

    private void applyWideDispatchQtyCellDisplay(
            SpreadsheetCell cell,
            WideRow wr,
            int dateIdx,
            List<LocalDate> axis,
            Map<String, DispatchQtyShortfallRow> wideShortfallMerged) {
        final double eps = 1e-3;
        double planAmt = wr.getAmount(dateIdx);
        double actualAmt = wr.getActualAmount(dateIdx);
        double aladdinAmt = aladdinPlanQtyForWideRow(wr, axis.get(dateIdx));
        if (isWideDispatchShortfall(wr, dateIdx)) {
            LocalDate dCell = axis.get(dateIdx);
            String tid = wr.getStatic("依頼NO");
            String mach = wr.getStatic(ResultDispatchSchema.COL_MACHINE);
            String k = DispatchTrialShortages.wideShortfallKey(tid, mach, dCell.toString());
            DispatchQtyShortfallRow sf = wideShortfallMerged.get(k);
            if (sf != null) {
                String targetFmt = ResultDispatchNormalizer.formatQty(sf.targetM());
                String doneFmt = ResultDispatchNormalizer.formatQty(sf.doneM());
                String aladdinFmt =
                        aladdinAmt > eps ? ResultDispatchNormalizer.formatQty(aladdinAmt) : "";
                setDispatchQtyCellDisplay(
                        cell,
                        filterDispatchQtyDisplayText(
                                formatDispatchPlanActualQtyText(
                                        aladdinFmt,
                                        targetFmt,
                                        doneFmt,
                                        stage3PlanActualSingleLineDisplay(),
                                        DispatchPlanQtyLineLabel.STAGE3),
                                dateQtyLineFilter,
                                stage3PlanActualSingleLineDisplay()),
                        stage3PlanActualSingleLineDisplay());
                tagDispatchDateQtyShortfallCell(cell, dispatchDateQtyMultilineCell());
                return;
            }
        }
        double snapPlan = stage3TrialSnapPlanForCell(wr.profileMap(), axis.get(dateIdx));
        if (hasPipelineStage3PlanningApplied() && !docHasActualDispatchQtyColumn()) {
            boolean stage3Revised =
                    isStage3QtyRevisedAfterTrial(
                            wr.profileMap(), axis.get(dateIdx), planAmt, actualAmt, eps);
            applyDispatchPlanActualQtyCellDisplay(
                    cell,
                    aladdinAmt,
                    snapPlan,
                    planAmt,
                    true,
                    eps,
                    stage3PlanActualSingleLineDisplay(),
                    stage3Revised,
                    false,
                    0.0,
                    dateQtyLineFilter,
                    effectiveStage3PlanningVariant());
            tagDispatchDateQtyCell(cell, dispatchDateQtyMultilineCell(), false);
            return;
        }
        if (!docHasActualDispatchQtyColumn() && hasStage21TrialApplied()) {
            double baseline = stage21BaselinePlanForCell(wr.profileMap(), axis.get(dateIdx));
            double stage21Amt = stage21ActualForCell(wr.profileMap(), axis.get(dateIdx));
            List<Stage3QtyLineSlot> slots =
                    applyDateQtyLineFilterToSlots(
                            buildStage21PreStage3CompareLineSlots(
                                    aladdinAmt, baseline, stage21Amt, eps),
                            dateQtyLineFilter);
            setDispatchQtyCellDisplay(cell, slots, false);
            tagDispatchDateQtyCell(cell, true, Math.abs(baseline - stage21Amt) > eps);
            return;
        }
        // 段階3試行後のみ旧配分を (段階2後) 表示。段階2のみのとき古い snapshot で planSlidAway すると幽霊行が出る。
        boolean planSlidAway =
                docHasActualDispatchQtyColumn()
                        && snapPlan > eps
                        && planAmt <= eps
                        && actualAmt <= eps;
        boolean planMovedToDate =
                snapPlan <= eps && (planAmt > eps || actualAmt > eps) && docHasActualDispatchQtyColumn();
        if (planSlidAway || planMovedToDate) {
            boolean stage3Revised =
                    !planSlidAway
                            && isStage3QtyRevisedAfterTrial(
                                    wr.profileMap(), axis.get(dateIdx), planAmt, actualAmt, eps);
            boolean stage21CompareSlide = hasStage21TrialApplied();
            double stage3BaselineSlide =
                    stage21CompareSlide
                            ? stage21BaselinePlanForCell(wr.profileMap(), axis.get(dateIdx))
                            : 0.0;
            applyDispatchPlanActualQtyCellDisplayWithPlanSlide(
                    cell,
                    aladdinAmt,
                    snapPlan,
                    planAmt,
                    actualAmt,
                    eps,
                    stage3PlanActualSingleLineDisplay(),
                    stage3Revised,
                    planSlidAway,
                    planMovedToDate,
                    stage21CompareSlide,
                    stage3BaselineSlide,
                    dateQtyLineFilter,
                    effectiveStage3PlanningVariant());
            tagDispatchDateQtyCell(
                    cell,
                    dispatchDateQtyMultilineCell(),
                    stage21CompareSlide
                            && Math.abs(stage3BaselineSlide - actualAmt) > eps);
            return;
        }
        boolean stage3Revised =
                isStage3QtyRevisedAfterTrial(
                        wr.profileMap(), axis.get(dateIdx), planAmt, actualAmt, eps);
        boolean appendStage21 = hasStage21TrialApplied();
        double stage21Amt =
                appendStage21 ? stage21ActualForCell(wr.profileMap(), axis.get(dateIdx)) : 0.0;
        double stage2Baseline =
                appendStage21 ? stage21BaselinePlanForCell(wr.profileMap(), axis.get(dateIdx)) : 0.0;
        applyDispatchPlanActualQtyCellDisplay(
                cell,
                aladdinAmt,
                planAmt,
                actualAmt,
                docHasActualDispatchQtyColumn(),
                eps,
                stage3PlanActualSingleLineDisplay(),
                stage3Revised,
                appendStage21,
                stage21Amt,
                dateQtyLineFilter,
                effectiveStage3PlanningVariant());
        tagDispatchDateQtyCell(
                cell,
                dispatchDateQtyMultilineCell(),
                appendStage21 && Math.abs(stage2Baseline - stage21Amt) > eps);
    }

    private void applyByDayDispatchQtyCellDisplay(
            SpreadsheetCell cell, ByDayRow br, int dateIdx, List<LocalDate> axis) {
        double aladdinAmt = aladdinPlanQtyForByDayRow(br, axis.get(dateIdx));
        double planAmt = br.getAmount(dateIdx);
        double actualAmt = br.getActualAmount(dateIdx);
        double eps = 1e-3;
        boolean stage3Revised =
                docHasActualDispatchQtyColumn() && Math.abs(planAmt - actualAmt) > eps;
        boolean appendStage21 = hasStage21TrialApplied();
        double stage21Amt =
                appendStage21 ? stage21ActualForByDay(br, axis.get(dateIdx)) : 0.0;
        double stage2Baseline =
                appendStage21 ? stage21BaselinePlanForByDay(br, axis.get(dateIdx)) : 0.0;
        if (!docHasActualDispatchQtyColumn() && appendStage21) {
            List<Stage3QtyLineSlot> slots =
                    applyDateQtyLineFilterToSlots(
                            buildStage21PreStage3CompareLineSlots(
                                    aladdinAmt, stage2Baseline, stage21Amt, eps),
                            dateQtyLineFilter);
            setDispatchQtyCellDisplay(cell, slots, false);
            tagDispatchDateQtyCell(cell, true, Math.abs(stage2Baseline - stage21Amt) > eps);
            return;
        }
        applyDispatchPlanActualQtyCellDisplay(
                cell,
                aladdinAmt,
                planAmt,
                actualAmt,
                docHasActualDispatchQtyColumn(),
                eps,
                stage3PlanActualSingleLineDisplay(),
                stage3Revised,
                appendStage21,
                stage21Amt,
                dateQtyLineFilter,
                effectiveStage3PlanningVariant());
        tagDispatchDateQtyCell(
                cell,
                dispatchDateQtyMultilineCell(),
                appendStage21 && Math.abs(stage2Baseline - stage21Amt) > eps);
    }

    private void scheduleStage3TrialPlanQtySnapshotCapture() {
        pendingStage3TrialSnapshotCapture = true;
    }

    private void clearStage3TrialPlanQtySnapshot() {
        stage3TrialPlanQtySnapshot.clear();
        pendingStage3TrialSnapshotCapture = false;
    }

    /** 段階2再実行時: 古い段階3 sidecar と比較用スナップショットを破棄し、表示を段階2モードへ戻す。 */
    private void clearStage3PlanningMeta(Path dispatchJsonPath) {
        clearStage3TrialPlanQtySnapshot();
        stage3PlanningVariant = ResultDispatchStage3Support.Stage3PlanningVariant.NONE;
        if (dispatchJsonPath != null) {
            Stage3PlanningMetaStore.deleteSidecar(dispatchJsonPath);
        }
        refreshDispatchPlanningStageBadge(dispatchJsonPath);
    }

    private boolean hasStage21TrialApplied() {
        return stage21TrialApplied && !stage21BaselinePlanQtySnapshot.isEmpty();
    }

    private void clearStage21TrialSnapshot(Path dispatchJsonPath) {
        stage21TrialApplied = false;
        stage21TrialMeta = Stage21TrialSnapshotStore.Stage21TrialMeta.empty();
        stage21BaselinePlanQtySnapshot.clear();
        stage21TimelineCalendarMeters = DispatchTimelineCalendarMetersIndex.empty();
        if (dispatchJsonPath != null) {
            Stage21TrialSnapshotStore.deleteSidecar(dispatchJsonPath);
        }
        refreshStage21AttendanceApplyPanel(dispatchJsonPath);
    }

    private void loadStage21TrialSnapshotFromDiskIfNeeded(Path dispatchJsonPath) {
        if (dispatchJsonPath == null) {
            return;
        }
        Stage21TrialSnapshotStore.Stage21TrialMeta meta =
                Stage21TrialSnapshotStore.tryLoadMeta(dispatchJsonPath);
        stage21TrialMeta = meta;
        if (meta.hasComparisonBaseline()) {
            stage21TrialApplied = true;
            stage21BaselinePlanQtySnapshot.clear();
            stage21BaselinePlanQtySnapshot.putAll(meta.entries());
        } else {
            stage21TrialApplied = false;
            stage21BaselinePlanQtySnapshot.clear();
            stage21TimelineCalendarMeters = DispatchTimelineCalendarMetersIndex.empty();
        }
        if (meta.hasPromotedToMain() || meta.hasComparisonBaseline()) {
            loadStage21TimelineFromDisk();
        }
        refreshStage21AttendanceApplyPanel(dispatchJsonPath);
        refreshDispatchPlanningStageBadge(dispatchJsonPath);
    }

    private void refreshDispatchPlanningStageBadge(Path dispatchJsonPath) {
        if (dispatchPlanningStageBadgeLabel == null) {
            return;
        }
        Path jsonPath =
                dispatchJsonPath != null
                        ? dispatchJsonPath
                        : (shell != null
                                ? AppPaths.resolveResultDispatchTableJsonPath(shell.snapshotUiEnv())
                                : null);
        stage3PlanningVariant = Stage3PlanningMetaStore.readPlanningVariant(jsonPath);
        ResultDispatchStage3Support.PlanningStage stage =
                ResultDispatchStage3Support.detectPlanningStage(jsonPath);
        ResultDispatchStage3Support.applyPlanningStageBadge(
                dispatchPlanningStageBadgeLabel,
                stage,
                stage == ResultDispatchStage3Support.PlanningStage.STAGE3
                        ? effectiveStage3PlanningVariant()
                        : ResultDispatchStage3Support.Stage3PlanningVariant.NONE);
    }

    private void captureStage21BaselineFromDocument(
            ResultDispatchDocument sourceDoc,
            List<LocalDate> axis,
            Path dispatchJsonPath,
            Path overtimeSimulationJson,
            Path stage21ResultDispatchJson) {
        stage21BaselinePlanQtySnapshot.clear();
        stage21BaselinePlanQtySnapshot.putAll(
                Stage21TrialSnapshotStore.captureStage2PlanFromDocument(sourceDoc, axis));
        Stage21TrialSnapshotStore.OverrideSummary summary =
                OvertimeSimulationOverridesReader.summarize(overtimeSimulationJson);
        stage21TrialApplied = true;
        stage21TrialMeta =
                new Stage21TrialSnapshotStore.Stage21TrialMeta(
                        true,
                        false,
                        stage21ResultDispatchJson != null
                                ? stage21ResultDispatchJson.toAbsolutePath().normalize().toString()
                                : "",
                        overtimeSimulationJson != null
                                ? overtimeSimulationJson.toAbsolutePath().normalize().toString()
                                : "",
                        summary,
                        java.time.LocalDateTime.now()
                                .format(java.time.format.DateTimeFormatter.ISO_LOCAL_DATE_TIME),
                        Map.copyOf(stage21BaselinePlanQtySnapshot));
        if (dispatchJsonPath != null) {
            Stage21TrialSnapshotStore.writeWithMeta(
                    dispatchJsonPath,
                    stage21BaselinePlanQtySnapshot,
                    stage21ResultDispatchJson,
                    overtimeSimulationJson,
                    summary);
        }
        refreshStage21AttendanceApplyPanel(dispatchJsonPath);
    }

    static String formatStage21CompletionStatus(
            Stage21TrialSnapshotStore.Stage21TrialMeta meta) {
        int n =
                meta != null && meta.overrideSummary() != null
                        ? meta.overrideSummary().totalChanges()
                        : 0;
        return "段階2.1 完了（勤怠 " + n + " 件適用）";
    }

    private void refreshStage21AttendanceApplyPanel(Path dispatchJsonPath) {
        Stage21AttendanceApplyStatusPanel.ViewModel vm =
                Stage21AttendanceApplyStatusPanel.build(
                        stage21TrialMeta, dispatchJsonPath, lastShortagesNote);
        Stage21AttendanceApplyStatusPanel.apply(
                stage21AttendanceApplyPanel,
                stage21AttendanceHeadlineLabel,
                stage21AttendanceSummaryLabel,
                stage21AttendanceOverridesLabel,
                stage21AttendancePythonLabel,
                stage21AttendanceAppliedAtLabel,
                vm);
    }

    private double stage21BaselinePlanForCell(Map<String, String> profile, LocalDate day) {
        if (profile == null || day == null) {
            return 0.0;
        }
        String key = Stage21TrialSnapshotStore.cellKey(profile, day);
        Double snap = stage21BaselinePlanQtySnapshot.get(key);
        return snap != null ? snap : 0.0;
    }

    private double stage21ActualForCell(Map<String, String> profile, LocalDate day) {
        if (profile == null || day == null || !stage21TimelineCalendarMeters.isLoaded()) {
            return 0.0;
        }
        Optional<Double> tl =
                stage21TimelineCalendarMeters.metersForTaskProfile(
                        profile.get("依頼NO"),
                        profile.get(ResultDispatchSchema.COL_PROCESS),
                        profile.get(ResultDispatchSchema.COL_MACHINE),
                        day);
        return tl.orElse(0.0);
    }

    private double stage21ActualForByDay(ByDayRow br, LocalDate day) {
        if (br == null || day == null || !stage21TimelineCalendarMeters.isLoaded()) {
            return 0.0;
        }
        Optional<Double> tl =
                stage21TimelineCalendarMeters.metersForProcessMachine(
                        br.process(), br.machine(), day);
        return tl.orElse(0.0);
    }

    /** 工程+機械×日ビュー: 同一機械・暦日のタスク別 baseline を合算する。 */
    private double stage21BaselinePlanForByDay(ByDayRow br, LocalDate day) {
        if (!hasStage21TrialApplied() || br == null || day == null) {
            return 0.0;
        }
        String machine = br.machine() != null ? br.machine().strip() : "";
        String dateIso = day.toString();
        double sum = 0.0;
        for (Map.Entry<String, Double> e : stage21BaselinePlanQtySnapshot.entrySet()) {
            String[] parts = e.getKey().split("\u0001", -1);
            if (parts.length >= 3
                    && machine.equals(parts[1].strip())
                    && dateIso.equals(parts[2].strip())) {
                sum += e.getValue();
            }
        }
        return sum;
    }

    private void captureStage3TrialPlanQtySnapshot(
            List<Map<String, String>> profiles, List<LocalDate> axis) {
        if (doc == null) {
            return;
        }
        captureStage3TrialPlanQtySnapshotFromDocument(doc, profiles, axis);
    }

    /** 配台試行開始前の doc から段階2後（当日配台）スナップショットを取る。 */
    private void captureStage3TrialPlanQtySnapshotFromDocument(
            ResultDispatchDocument sourceDoc, List<LocalDate> axis) {
        if (sourceDoc == null || axis == null || axis.isEmpty()) {
            return;
        }
        List<Map<String, String>> profiles =
                ResultDispatchPivot.distinctWideTaskProfiles(
                        sourceDoc.columns(),
                        sourceDoc.rows(),
                        ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
        captureStage3TrialPlanQtySnapshotFromDocument(sourceDoc, profiles, axis);
    }

    private void captureStage3TrialPlanQtySnapshotFromDocument(
            ResultDispatchDocument sourceDoc,
            List<Map<String, String>> profiles,
            List<LocalDate> axis) {
        stage3TrialPlanQtySnapshot.clear();
        pendingStage3TrialSnapshotCapture = false;
        if (sourceDoc == null || profiles == null || axis == null) {
            return;
        }
        for (Map<String, String> profile : profiles) {
            for (LocalDate day : axis) {
                double plan =
                        ResultDispatchPivot.sumQuantityForProfileAndDateForWideMerge(
                                sourceDoc.rows(),
                                profile,
                                day,
                                ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
                String key =
                        DispatchTrialShortages.wideShortfallKey(
                                profile.get("依頼NO"),
                                profile.get(ResultDispatchSchema.COL_MACHINE),
                                day.toString());
                stage3TrialPlanQtySnapshot.put(key, plan);
            }
        }
    }

    private List<LocalDate> snapshotDateAxisForTrialPlanQtyCapture(ResultDispatchDocument snapshotDoc) {
        if (!dateAxis.isEmpty()) {
            return new ArrayList<>(dateAxis);
        }
        if (snapshotDoc != null) {
            Map<String, Map<String, Map<String, Map<String, Double>>>> lookup =
                    aladdinPlanLookup.isEmpty()
                            ? loadAladdinPlanLookupForDisplay()
                            : aladdinPlanLookup;
            List<LocalDate> axis =
                    DispatchInteractiveDateAxis.computeInclusiveRange(
                            snapshotDoc, lookup, lastDispatchShortfallRows, dateAxisPastDays);
            if (!axis.isEmpty()) {
                return axis;
            }
        }
        return computeDateAxisList();
    }

    private double stage3TrialSnapPlanForCell(Map<String, String> profile, LocalDate day) {
        if (profile == null || day == null) {
            return 0.0;
        }
        String key =
                DispatchTrialShortages.wideShortfallKey(
                        profile.get("依頼NO"),
                        profile.get(ResultDispatchSchema.COL_MACHINE),
                        day.toString());
        Double snap = stage3TrialPlanQtySnapshot.get(key);
        return snap != null ? snap : 0.0;
    }

    /**
     * 試行後にユーザーが当日配台数量を手動変更したセルか。
     * 試行前に当該暦日の配台が無かった日（snap≈0）や、タイムライン未達（plan≠actual）は含めない。
     */
    private boolean isStage3QtyRevisedAfterTrial(
            Map<String, String> profile,
            LocalDate day,
            double planAmt,
            double actualAmt,
            double eps) {
        if ((!docHasActualDispatchQtyColumn() && !hasPipelineStage3PlanningApplied()) || profile == null || day == null) {
            return false;
        }
        String key =
                DispatchTrialShortages.wideShortfallKey(
                        profile.get("依頼NO"),
                        profile.get(ResultDispatchSchema.COL_MACHINE),
                        day.toString());
        Double snap = stage3TrialPlanQtySnapshot.get(key);
        if (snap == null || snap <= eps) {
            return false;
        }
        return Math.abs(planAmt - snap) > eps;
    }

    private double aladdinPlanQtyForWideRow(WideRow wr, LocalDate day) {
        if (aladdinPlanLookup.isEmpty() || day == null) {
            return 0.0;
        }
        return AladdinShapedPlanQtyLookup.lookup(
                aladdinPlanLookup,
                wr.getStatic(ResultDispatchSchema.COL_MACHINE),
                wr.getStatic("依頼NO"),
                day.format(ALADDIN_PLAN_DATE_FMT),
                wr.getStatic(ResultDispatchSchema.COL_PROCESS));
    }

    private double aladdinPlanQtyForByDayRow(ByDayRow br, LocalDate day) {
        if (aladdinPlanLookup.isEmpty() || day == null || doc == null) {
            return 0.0;
        }
        String dateStr = day.format(ALADDIN_PLAN_DATE_FMT);
        String process = br.process();
        String machine = br.machine();
        List<String> cols = doc.columns();
        double sum = 0.0;
        Set<String> seenTids = new HashSet<>();
        for (Map<String, String> row : doc.rows()) {
            if (!Objects.equals(process, row.get(ResultDispatchSchema.COL_PROCESS))
                    || !Objects.equals(machine, row.get(ResultDispatchSchema.COL_MACHINE))) {
                continue;
            }
            String tid = row.getOrDefault("依頼NO", "").strip();
            if (tid.isEmpty() || !seenTids.add(tid)) {
                continue;
            }
            sum +=
                    AladdinShapedPlanQtyLookup.lookup(
                            aladdinPlanLookup, machine, tid, dateStr, process);
        }
        return sum;
    }

    private boolean dispatchDateQtyMultilineCell() {
        if (stage3PlanActualSingleLineDisplay()) {
            return false;
        }
        return showsStage3QtyMultilineDisplay() || !aladdinPlanLookup.isEmpty();
    }
    /**
     * 段階3の2行表示: 固定行高＋{@code \\n}（{@code -fx-wrap-text} は使わない）。単行表示時は既定行高。
     */
    private void applyDispatchStage3QtyRowPresentation(GridBase grid) {
        if (grid == null) {
            return;
        }
        if (dispatchDateQtyMultilineCell()) {
            double rowHeightPx =
                    showsStage3QtyMultilineDisplay()
                            ? (hasStage21TrialApplied()
                                    ? DISPATCH_ALADDIN_STAGE21_MULTILINE_ROW_HEIGHT_PX
                                    : DISPATCH_ALADDIN_STAGE3_MULTILINE_ROW_HEIGHT_PX)
                            : DISPATCH_STAGE3_MULTILINE_ROW_HEIGHT_PX;
            SpreadsheetTabularSupport.applySpreadsheetGridRowHeightsAndWrap(
                    grid,
                    false,
                    100.0,
                    24.0,
                    rowHeightPx);
        } else {
            SpreadsheetTabularSupport.applySpreadsheetGridRowHeightsAndWrap(grid, false, 100.0);
        }
    }

    /**
     * 段階3試行後: (アラ計画)・編集目標（当日配台数量）・実績（実配台数量）を同一セルに表示する。
     * DnD・ダブルクリック編集は目標のみ更新する。
     */
    private static void applyDispatchPlanActualQtyCellDisplay(
            SpreadsheetCell cell,
            double aladdinPlanAmt,
            double planAmt,
            double actualAmt,
            boolean hasActualColumn,
            double eps,
            boolean singleLineDisplay,
            boolean stage3RevisedAfterTrial) {
        applyDispatchPlanActualQtyCellDisplay(
                cell,
                aladdinPlanAmt,
                planAmt,
                actualAmt,
                hasActualColumn,
                eps,
                singleLineDisplay,
                stage3RevisedAfterTrial,
                false,
                0.0,
                DispatchInteractiveDateQtyLineFilterPrefs.defaults());
    }

    private static void applyDispatchPlanActualQtyCellDisplay(
            SpreadsheetCell cell,
            double aladdinPlanAmt,
            double planAmt,
            double actualAmt,
            boolean hasActualColumn,
            double eps,
            boolean singleLineDisplay,
            boolean stage3RevisedAfterTrial,
            boolean stage21CompareMode,
            double stage3BaselineActualAmt,
            DispatchInteractiveDateQtyLineFilterPrefs lineFilter) {
        applyDispatchPlanActualQtyCellDisplay(
                cell,
                aladdinPlanAmt,
                planAmt,
                actualAmt,
                hasActualColumn,
                eps,
                singleLineDisplay,
                stage3RevisedAfterTrial,
                stage21CompareMode,
                stage3BaselineActualAmt,
                lineFilter,
                ResultDispatchStage3Support.Stage3PlanningVariant.LEGACY);
    }

    private static void applyDispatchPlanActualQtyCellDisplay(
            SpreadsheetCell cell,
            double aladdinPlanAmt,
            double planAmt,
            double actualAmt,
            boolean hasActualColumn,
            double eps,
            boolean singleLineDisplay,
            boolean stage3RevisedAfterTrial,
            boolean stage21CompareMode,
            double stage3BaselineActualAmt,
            DispatchInteractiveDateQtyLineFilterPrefs lineFilter,
            ResultDispatchStage3Support.Stage3PlanningVariant variant) {
        if (hasActualColumn && !singleLineDisplay) {
            List<Stage3QtyLineSlot> slots =
                    applyDateQtyLineFilterToSlots(
                            buildStage3QtyFixedLineSlots(
                                    aladdinPlanAmt,
                                    planAmt,
                                    actualAmt,
                                    stage3RevisedAfterTrial,
                                    eps,
                                    stage21CompareMode,
                                    stage3BaselineActualAmt,
                                    variant),
                            lineFilter);
            setDispatchQtyCellDisplay(cell, slots, false);
            return;
        }
        String qtxt =
                filterDispatchQtyDisplayText(
                        formatDispatchPlanActualQtyDisplay(
                                aladdinPlanAmt,
                                planAmt,
                                actualAmt,
                                hasActualColumn,
                                eps,
                                singleLineDisplay,
                                stage3RevisedAfterTrial,
                                stage21CompareMode,
                                stage3BaselineActualAmt,
                                variant),
                        lineFilter,
                        singleLineDisplay);
        setDispatchQtyCellDisplay(cell, qtxt, singleLineDisplay);
    }

    /** 段階3: 1=(アラ計画)、2=(段階3前)、3=(段階3後)または(段階3改)。 */
    private static void setDispatchQtyCellDisplay(
            SpreadsheetCell cell, List<Stage3QtyLineSlot> slots, boolean singleLineDisplay) {
        if (slots == null || slots.stream().noneMatch(Stage3QtyLineSlot::visible)) {
            clearDispatchQtyCellGraphic(cell);
            SpreadsheetTabularSupport.setSpreadsheetCellDisplayValue(cell, "");
            return;
        }
        if (singleLineDisplay) {
            setDispatchQtyCellDisplay(
                    cell, formatStage3FixedSlotsAsText(slots, true), true);
            return;
        }
        String copyText = formatStage3FixedSlotsAsText(slots, false);
        applyDispatchQtyGraphicCellDisplay(
                cell, buildStage3QtyFixedLineGraphic(slots), copyText);
    }

    /**
     * 日付セル表示。(段階3後) 行は {@link #DISPATCH_STAGE3_AFTER_LINE_STYLE_CLASS} で太字にする。
     */
    private static void setDispatchQtyCellDisplay(
            SpreadsheetCell cell, String qtxt, boolean singleLineDisplay) {
        boolean useStyledGraphic =
                qtxt != null
                        && !qtxt.isBlank()
                        && (ResultDispatchStage3Support.isStage3RevisedQtyLine(qtxt)
                                || qtxt.contains(LABEL_STAGE21_ACTUAL)
                                || ResultDispatchStage3Support.isStage3AfterQtyLine(qtxt));
        if (useStyledGraphic) {
            applyDispatchQtyGraphicCellDisplay(
                    cell,
                    buildDispatchPlanActualQtyGraphic(qtxt, singleLineDisplay),
                    qtxt != null ? qtxt : "");
        } else {
            clearDispatchQtyCellGraphic(cell);
            SpreadsheetTabularSupport.setSpreadsheetCellDisplayValue(cell, qtxt != null ? qtxt : "");
        }
    }

    static List<Stage3QtyLineSlot> buildStage21PreStage3CompareLineSlots(
            double aladdinPlanAmt, double stage2BaselineAmt, double stage21ActualAmt, double eps) {
        List<Stage3QtyLineSlot> slots = new ArrayList<>(STAGE21_PRE_STAGE3_QTY_LINE_COUNT);
        slots.add(stage3QtyLineSlot(LABEL_ALADDIN_PLAN, aladdinPlanAmt, eps));
        slots.add(stage3QtyLineSlot(LABEL_STAGE2_PLAN, stage2BaselineAmt, eps));
        slots.add(stage3QtyLineSlot(LABEL_STAGE21_ACTUAL, stage21ActualAmt, eps));
        return slots;
    }

    static List<Stage3QtyLineSlot> buildStage3QtyFixedLineSlots(
            double aladdinPlanAmt,
            double planAmt,
            double actualAmt,
            boolean stage3RevisedAfterTrial,
            double eps) {
        return buildStage3QtyFixedLineSlots(
                aladdinPlanAmt, planAmt, actualAmt, stage3RevisedAfterTrial, eps, false, 0.0);
    }

    static List<Stage3QtyLineSlot> buildStage3QtyFixedLineSlots(
            double aladdinPlanAmt,
            double planAmt,
            double actualAmt,
            boolean stage3RevisedAfterTrial,
            double eps,
            boolean appendStage21Line,
            double stage21ActualAmt) {
        return buildStage3QtyFixedLineSlots(
                aladdinPlanAmt,
                planAmt,
                actualAmt,
                stage3RevisedAfterTrial,
                eps,
                appendStage21Line,
                stage21ActualAmt,
                ResultDispatchStage3Support.Stage3PlanningVariant.LEGACY);
    }

    static List<Stage3QtyLineSlot> buildStage3QtyFixedLineSlots(
            double aladdinPlanAmt,
            double planAmt,
            double actualAmt,
            boolean stage3RevisedAfterTrial,
            double eps,
            boolean appendStage21Line,
            double stage21ActualAmt,
            ResultDispatchStage3Support.Stage3PlanningVariant variant) {
        ResultDispatchStage3Support.Stage3PlanningVariant v =
                variant != null
                        ? variant
                        : ResultDispatchStage3Support.Stage3PlanningVariant.LEGACY;
        String actualLabel = v.actualQtyLabel();
        String revisedLabel = v.revisedQtyLabel();
        int lineCount =
                appendStage21Line ? STAGE21_QTY_FIXED_LINE_COUNT : STAGE3_QTY_FIXED_LINE_COUNT;
        List<Stage3QtyLineSlot> slots = new ArrayList<>(lineCount);
        slots.add(stage3QtyLineSlot(LABEL_ALADDIN_PLAN, aladdinPlanAmt, eps));
        if (stage3RevisedAfterTrial) {
            slots.add(stage3QtyEmptyLineSlot());
            slots.add(stage3QtyLineSlot(revisedLabel, planAmt, eps));
        } else {
            slots.add(stage3QtyLineSlot(LABEL_STAGE3_PLAN, planAmt, eps));
            slots.add(stage3QtyLineSlot(actualLabel, actualAmt, eps));
        }
        if (appendStage21Line) {
            while (slots.size() < STAGE21_QTY_FIXED_LINE_COUNT - 1) {
                slots.add(stage3QtyEmptyLineSlot());
            }
            slots.add(stage3QtyLineSlot(LABEL_STAGE21_ACTUAL, stage21ActualAmt, eps));
        }
        while (slots.size() < lineCount) {
            slots.add(stage3QtyEmptyLineSlot());
        }
        return slots;
    }

    /**
     * 配台日スライド後: 旧日付は (段階2後)、新日付は (段階3後) のみ（(段階3前) は出さない）。
     */
    static List<Stage3QtyLineSlot> buildStage3QtyFixedLineSlotsWithPlanSlide(
            double aladdinPlanAmt,
            double snapPlanAmt,
            double planAmt,
            double actualAmt,
            boolean stage3RevisedAfterTrial,
            boolean planSlidAway,
            boolean planMovedToDate,
            double eps) {
        return buildStage3QtyFixedLineSlotsWithPlanSlide(
                aladdinPlanAmt,
                snapPlanAmt,
                planAmt,
                actualAmt,
                stage3RevisedAfterTrial,
                planSlidAway,
                planMovedToDate,
                eps,
                false,
                0.0);
    }

    static List<Stage3QtyLineSlot> buildStage3QtyFixedLineSlotsWithPlanSlide(
            double aladdinPlanAmt,
            double snapPlanAmt,
            double planAmt,
            double actualAmt,
            boolean stage3RevisedAfterTrial,
            boolean planSlidAway,
            boolean planMovedToDate,
            double eps,
            boolean stage21CompareMode,
            double stage3BaselineActualAmt) {
        return buildStage3QtyFixedLineSlotsWithPlanSlide(
                aladdinPlanAmt,
                snapPlanAmt,
                planAmt,
                actualAmt,
                stage3RevisedAfterTrial,
                planSlidAway,
                planMovedToDate,
                eps,
                stage21CompareMode,
                stage3BaselineActualAmt,
                ResultDispatchStage3Support.Stage3PlanningVariant.LEGACY);
    }

    static List<Stage3QtyLineSlot> buildStage3QtyFixedLineSlotsWithPlanSlide(
            double aladdinPlanAmt,
            double snapPlanAmt,
            double planAmt,
            double actualAmt,
            boolean stage3RevisedAfterTrial,
            boolean planSlidAway,
            boolean planMovedToDate,
            double eps,
            boolean stage21CompareMode,
            double stage3BaselineActualAmt,
            ResultDispatchStage3Support.Stage3PlanningVariant variant) {
        ResultDispatchStage3Support.Stage3PlanningVariant v =
                variant != null
                        ? variant
                        : ResultDispatchStage3Support.Stage3PlanningVariant.LEGACY;
        String actualLabel = v.actualQtyLabel();
        String revisedLabel = v.revisedQtyLabel();
        int lineCount = stage21CompareMode ? STAGE21_QTY_FIXED_LINE_COUNT : STAGE3_QTY_FIXED_LINE_COUNT;
        List<Stage3QtyLineSlot> slots = new ArrayList<>(lineCount);
        slots.add(stage3QtyLineSlot(LABEL_ALADDIN_PLAN, aladdinPlanAmt, eps));
        if (planSlidAway) {
            slots.add(stage3QtyLineSlot(LABEL_STAGE2_PLAN, snapPlanAmt, eps));
            slots.add(stage3QtyEmptyLineSlot());
            while (slots.size() < lineCount) {
                slots.add(stage3QtyEmptyLineSlot());
            }
            return slots;
        }
        if (planMovedToDate) {
            if (stage3RevisedAfterTrial) {
                if (stage21CompareMode) {
                    slots.add(stage3QtyLineSlot(revisedLabel, planAmt, eps));
                    slots.add(stage3QtyLineSlot(actualLabel, stage3BaselineActualAmt, eps));
                    slots.add(stage3QtyLineSlot(LABEL_STAGE21_ACTUAL, actualAmt, eps));
                } else {
                    slots.add(stage3QtyEmptyLineSlot());
                    slots.add(stage3QtyLineSlot(revisedLabel, planAmt, eps));
                }
            } else {
                slots.add(stage3QtyEmptyLineSlot());
                double afterAmt = actualAmt > eps ? actualAmt : planAmt;
                if (stage21CompareMode) {
                    slots.add(stage3QtyLineSlot(actualLabel, stage3BaselineActualAmt, eps));
                    slots.add(stage3QtyLineSlot(LABEL_STAGE21_ACTUAL, afterAmt, eps));
                } else {
                    slots.add(stage3QtyLineSlot(actualLabel, afterAmt, eps));
                }
            }
            while (slots.size() < lineCount) {
                slots.add(stage3QtyEmptyLineSlot());
            }
            return slots;
        }
        return buildStage3QtyFixedLineSlots(
                aladdinPlanAmt,
                planAmt,
                actualAmt,
                stage3RevisedAfterTrial,
                eps,
                stage21CompareMode,
                stage3BaselineActualAmt,
                v);
    }

    private static void applyDispatchPlanActualQtyCellDisplayWithPlanSlide(
            SpreadsheetCell cell,
            double aladdinPlanAmt,
            double snapPlanAmt,
            double planAmt,
            double actualAmt,
            double eps,
            boolean singleLineDisplay,
            boolean stage3RevisedAfterTrial,
            boolean planSlidAway,
            boolean planMovedToDate,
            boolean stage21CompareMode,
            double stage3BaselineActualAmt,
            DispatchInteractiveDateQtyLineFilterPrefs lineFilter) {
        applyDispatchPlanActualQtyCellDisplayWithPlanSlide(
                cell,
                aladdinPlanAmt,
                snapPlanAmt,
                planAmt,
                actualAmt,
                eps,
                singleLineDisplay,
                stage3RevisedAfterTrial,
                planSlidAway,
                planMovedToDate,
                stage21CompareMode,
                stage3BaselineActualAmt,
                lineFilter,
                ResultDispatchStage3Support.Stage3PlanningVariant.LEGACY);
    }

    private static void applyDispatchPlanActualQtyCellDisplayWithPlanSlide(
            SpreadsheetCell cell,
            double aladdinPlanAmt,
            double snapPlanAmt,
            double planAmt,
            double actualAmt,
            double eps,
            boolean singleLineDisplay,
            boolean stage3RevisedAfterTrial,
            boolean planSlidAway,
            boolean planMovedToDate,
            boolean stage21CompareMode,
            double stage3BaselineActualAmt,
            DispatchInteractiveDateQtyLineFilterPrefs lineFilter,
            ResultDispatchStage3Support.Stage3PlanningVariant variant) {
        if (!singleLineDisplay) {
            List<Stage3QtyLineSlot> slots =
                    applyDateQtyLineFilterToSlots(
                            buildStage3QtyFixedLineSlotsWithPlanSlide(
                                    aladdinPlanAmt,
                                    snapPlanAmt,
                                    planAmt,
                                    actualAmt,
                                    stage3RevisedAfterTrial,
                                    planSlidAway,
                                    planMovedToDate,
                                    eps,
                                    stage21CompareMode,
                                    stage3BaselineActualAmt,
                                    variant),
                            lineFilter);
            setDispatchQtyCellDisplay(cell, slots, false);
            return;
        }
        String qtxt =
                filterDispatchQtyDisplayText(
                        formatStage3FixedSlotsAsText(
                                buildStage3QtyFixedLineSlotsWithPlanSlide(
                                        aladdinPlanAmt,
                                        snapPlanAmt,
                                        planAmt,
                                        actualAmt,
                                        stage3RevisedAfterTrial,
                                        planSlidAway,
                                        planMovedToDate,
                                        eps,
                                        stage21CompareMode,
                                        stage3BaselineActualAmt,
                                        variant),
                                true),
                        lineFilter,
                        true);
        setDispatchQtyCellDisplay(cell, qtxt, true);
    }

    static List<Stage3QtyLineSlot> applyDateQtyLineFilterToSlots(
            List<Stage3QtyLineSlot> slots, DispatchInteractiveDateQtyLineFilterPrefs filter) {
        if (slots == null || slots.isEmpty()) {
            return slots;
        }
        if (filter == null || dateQtyLineFilterShowsAll(slots, filter)) {
            return slots;
        }
        List<Stage3QtyLineSlot> out = new ArrayList<>(slots.size());
        for (Stage3QtyLineSlot slot : slots) {
            if (!slot.visible()) {
                out.add(slot);
                continue;
            }
            String line = slot.lineText();
            if (line.startsWith(LABEL_ALADDIN_PLAN)) {
                out.add(filter.showAladdinPlan() ? slot : stage3QtyEmptyLineSlot());
            } else if (line.startsWith(LABEL_STAGE2_PLAN)) {
                out.add(filter.showStage2Plan() ? slot : stage3QtyEmptyLineSlot());
            } else if (line.startsWith(LABEL_STAGE3_PLAN)) {
                out.add(filter.showStage3Plan() ? slot : stage3QtyEmptyLineSlot());
            } else if (line.startsWith(LABEL_STAGE21_ACTUAL)) {
                out.add(filter.showStage21After() ? slot : stage3QtyEmptyLineSlot());
            } else if (ResultDispatchStage3Support.isStage3AfterQtyLine(line)
                    || ResultDispatchStage3Support.isStage3RevisedQtyLine(line)) {
                out.add(filter.showStage3After() ? slot : stage3QtyEmptyLineSlot());
            } else {
                out.add(slot);
            }
        }
        return out;
    }

    private static boolean dateQtyLineFilterShowsAll(
            List<Stage3QtyLineSlot> slots, DispatchInteractiveDateQtyLineFilterPrefs filter) {
        if (!filter.showAladdinPlan() || !filter.showStage2Plan() || !filter.showStage3Plan() || !filter.showStage3After()) {
            return false;
        }
        boolean hasStage21Line =
                slots != null
                        && slots.stream()
                                .anyMatch(
                                        s ->
                                                s.visible()
                                                        && s.lineText()
                                                                .startsWith(LABEL_STAGE21_ACTUAL));
        return !hasStage21Line || filter.showStage21After();
    }

    static String filterDispatchQtyDisplayText(
            String text,
            DispatchInteractiveDateQtyLineFilterPrefs filter,
            boolean singleLineDisplay) {
        if (text == null || text.isBlank() || filter == null) {
            return text != null ? text : "";
        }
        if (filter.showAladdinPlan()
                && filter.showStage2Plan()
                && filter.showStage3Plan()
                && filter.showStage3After()
                && (!text.contains(LABEL_STAGE21_ACTUAL) || filter.showStage21After())) {
            return text;
        }
        if (singleLineDisplay) {
            StringBuilder sb = new StringBuilder();
            for (String part : text.split(" ", -1)) {
                if (part.isEmpty() || !dispatchQtyLineVisible(part, filter)) {
                    continue;
                }
                if (!sb.isEmpty()) {
                    sb.append(' ');
                }
                sb.append(part);
            }
            return sb.toString();
        }
        StringBuilder sb = new StringBuilder();
        for (String line : text.split("\n", -1)) {
            if (line.isEmpty() || !dispatchQtyLineVisible(line, filter)) {
                continue;
            }
            if (!sb.isEmpty()) {
                sb.append('\n');
            }
            sb.append(line);
        }
        return sb.toString();
    }

    private static boolean dispatchQtyLineVisible(
            String line, DispatchInteractiveDateQtyLineFilterPrefs filter) {
        if (line.startsWith(LABEL_ALADDIN_PLAN)) {
            return filter.showAladdinPlan();
        }
        if (line.startsWith(LABEL_STAGE2_PLAN)) {
            return filter.showStage2Plan();
        }
        if (line.startsWith(LABEL_STAGE3_PLAN)) {
            return filter.showStage3Plan();
        }
        if (line.startsWith(LABEL_STAGE21_ACTUAL)) {
            return filter.showStage21After();
        }
        if (ResultDispatchStage3Support.isStage3AfterQtyLine(line)
                || ResultDispatchStage3Support.isStage3RevisedQtyLine(line)) {
            return filter.showStage3After();
        }
        return true;
    }

    private static Stage3QtyLineSlot stage3QtyLineSlot(String label, double amt, double eps) {
        if (amt <= eps) {
            return stage3QtyEmptyLineSlot();
        }
        return new Stage3QtyLineSlot(label + ResultDispatchNormalizer.formatQty(amt), true);
    }

    private static Stage3QtyLineSlot stage3QtyEmptyLineSlot() {
        return new Stage3QtyLineSlot("", false);
    }

    static String formatStage3FixedSlotsAsText(List<Stage3QtyLineSlot> slots, boolean singleLineDisplay) {
        if (slots == null || slots.stream().noneMatch(Stage3QtyLineSlot::visible)) {
            return "";
        }
        if (singleLineDisplay) {
            StringBuilder sb = new StringBuilder();
            for (Stage3QtyLineSlot slot : slots) {
                if (!slot.visible()) {
                    continue;
                }
                if (!sb.isEmpty()) {
                    sb.append(' ');
                }
                sb.append(slot.lineText());
            }
            return sb.toString();
        }
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < slots.size(); i++) {
            if (i > 0) {
                sb.append('\n');
            }
            Stage3QtyLineSlot slot = slots.get(i);
            sb.append(slot.visible() ? slot.lineText() : "");
        }
        return sb.toString();
    }

    private static Node buildStage3QtyFixedLineGraphic(List<Stage3QtyLineSlot> slots) {
        VBox box = new VBox(0);
        box.setPadding(new Insets(2, 4, 2, 4));
        box.setFillWidth(true);
        for (Stage3QtyLineSlot slot : slots) {
            Label lbl = new Label(slot.visible() ? slot.lineText() : "\u00a0");
            lbl.getStyleClass().add("delivery-calendar-triple-slot");
            if (slot.visible()) {
                applyDispatchStage3QtyLineLabelStyle(lbl, slot.lineText());
            } else {
                lbl.getStyleClass().add("delivery-calendar-triple-slot-empty");
            }
            lbl.setWrapText(false);
            lbl.setMaxWidth(Double.MAX_VALUE);
            box.getChildren().add(lbl);
        }
        return box;
    }

    private static Node buildDispatchPlanActualQtyGraphic(String text, boolean singleLineDisplay) {
        if (singleLineDisplay) {
            return buildDispatchPlanActualQtySingleLineGraphic(text);
        }
        VBox box = new VBox(0);
        box.setPadding(new Insets(2, 4, 2, 4));
        box.setFillWidth(true);
        for (String line : text.split("\n", -1)) {
            if (line.isEmpty()) {
                continue;
            }
            Label lbl = new Label(line);
            lbl.getStyleClass().add("delivery-calendar-triple-slot");
            applyDispatchStage3QtyLineLabelStyle(lbl, line);
            lbl.setWrapText(false);
            lbl.setMaxWidth(Double.MAX_VALUE);
            box.getChildren().add(lbl);
        }
        return box;
    }

    private static Node buildDispatchPlanActualQtySingleLineGraphic(String text) {
        TextFlow flow = new TextFlow();
        flow.setPadding(new Insets(2, 4, 2, 4));
        int idx = 0;
        while (idx < text.length()) {
            int next = indexOfNextDispatchQtyStyledSegment(text, idx);
            if (next < 0) {
                Text tail = new Text(text.substring(idx));
                tail.setStyle(STAGE3_QTY_DEFAULT_TEXT_INLINE_STYLE);
                flow.getChildren().add(tail);
                break;
            }
            if (next > idx) {
                Text plain = new Text(text.substring(idx, next));
                plain.setStyle(STAGE3_QTY_DEFAULT_TEXT_INLINE_STYLE);
                flow.getChildren().add(plain);
            }
            int labelLen = dispatchQtyStyledSegmentLabelLength(text, next);
            int end = segmentEndIndex(text, next + labelLen);
            Text styled = new Text(text.substring(next, end));
            styled.setStyle(dispatchQtyStyledSegmentInlineStyle(text, next));
            flow.getChildren().add(styled);
            idx = end;
        }
        return flow;
    }

    private static int indexOfNextDispatchQtyStyledSegment(String text, int from) {
        int next = -1;
        next = minIfEarlier(next, text.indexOf(LABEL_STAGE21_ACTUAL, from));
        for (ResultDispatchStage3Support.Stage3PlanningVariant v :
                ResultDispatchStage3Support.Stage3PlanningVariant.values()) {
            if (v == ResultDispatchStage3Support.Stage3PlanningVariant.NONE) {
                continue;
            }
            next = minIfEarlier(next, text.indexOf(v.revisedQtyLabel(), from));
            next = minIfEarlier(next, text.indexOf(v.actualQtyLabel(), from));
        }
        return next;
    }

    private static int minIfEarlier(int current, int candidate) {
        if (candidate < 0) {
            return current;
        }
        return current < 0 || candidate < current ? candidate : current;
    }

    private static int dispatchQtyStyledSegmentLabelLength(String text, int index) {
        if (text.startsWith(LABEL_STAGE21_ACTUAL, index)) {
            return LABEL_STAGE21_ACTUAL.length();
        }
        String rest = text.substring(index);
        if (ResultDispatchStage3Support.isStage3RevisedQtyLine(rest)) {
            for (ResultDispatchStage3Support.Stage3PlanningVariant v :
                    ResultDispatchStage3Support.Stage3PlanningVariant.values()) {
                if (v == ResultDispatchStage3Support.Stage3PlanningVariant.NONE) {
                    continue;
                }
                String label = v.revisedQtyLabel();
                if (rest.startsWith(label)) {
                    return label.length();
                }
            }
        }
        if (ResultDispatchStage3Support.isStage3AfterQtyLine(rest)) {
            for (ResultDispatchStage3Support.Stage3PlanningVariant v :
                    ResultDispatchStage3Support.Stage3PlanningVariant.values()) {
                if (v == ResultDispatchStage3Support.Stage3PlanningVariant.NONE) {
                    continue;
                }
                String label = v.actualQtyLabel();
                if (rest.startsWith(label)) {
                    return label.length();
                }
            }
        }
        return 0;
    }

    private static String dispatchQtyStyledSegmentInlineStyle(String text, int index) {
        String rest = text.substring(index);
        if (ResultDispatchStage3Support.isStage3RevisedQtyLine(rest)) {
            return "-fx-font-weight: bold; -fx-fill: #1565C0;";
        }
        if (rest.startsWith(LABEL_STAGE21_ACTUAL)) {
            return "-fx-font-weight: bold; -fx-fill: #2E7D32;";
        }
        if (ResultDispatchStage3Support.isStage3AfterQtyLine(rest)) {
            return "-fx-font-weight: bold; -fx-fill: #111111;";
        }
        return STAGE3_QTY_DEFAULT_TEXT_INLINE_STYLE;
    }

    private static int segmentEndIndex(String text, int qtyStart) {
        int nextOpen = text.indexOf(" (", qtyStart);
        return nextOpen >= 0 ? nextOpen : text.length();
    }

    /** (アラ計画)・目標 m（段階3前）・実績 m（段階3後）をラベル付きで表示する。 */
    static String formatDispatchPlanActualQtyDisplay(
            double aladdinPlanAmt,
            double planAmt,
            double actualAmt,
            boolean hasActualColumn,
            double eps,
            boolean singleLineDisplay) {
        return formatDispatchPlanActualQtyDisplay(
                aladdinPlanAmt,
                planAmt,
                actualAmt,
                hasActualColumn,
                eps,
                singleLineDisplay,
                false);
    }

    static String formatDispatchPlanActualQtyDisplay(
            double aladdinPlanAmt,
            double planAmt,
            double actualAmt,
            boolean hasActualColumn,
            double eps,
            boolean singleLineDisplay,
            boolean stage3RevisedAfterTrial) {
        return formatDispatchPlanActualQtyDisplay(
                aladdinPlanAmt,
                planAmt,
                actualAmt,
                hasActualColumn,
                eps,
                singleLineDisplay,
                stage3RevisedAfterTrial,
                false,
                0.0);
    }

    static String formatDispatchPlanActualQtyDisplay(
            double aladdinPlanAmt,
            double planAmt,
            double actualAmt,
            boolean hasActualColumn,
            double eps,
            boolean singleLineDisplay,
            boolean stage3RevisedAfterTrial,
            boolean stage21CompareMode,
            double stage3BaselineActualAmt) {
        return formatDispatchPlanActualQtyDisplay(
                aladdinPlanAmt,
                planAmt,
                actualAmt,
                hasActualColumn,
                eps,
                singleLineDisplay,
                stage3RevisedAfterTrial,
                stage21CompareMode,
                stage3BaselineActualAmt,
                ResultDispatchStage3Support.Stage3PlanningVariant.LEGACY);
    }

    static String formatDispatchPlanActualQtyDisplay(
            double aladdinPlanAmt,
            double planAmt,
            double actualAmt,
            boolean hasActualColumn,
            double eps,
            boolean singleLineDisplay,
            boolean stage3RevisedAfterTrial,
            boolean stage21CompareMode,
            double stage3BaselineActualAmt,
            ResultDispatchStage3Support.Stage3PlanningVariant variant) {
        boolean hasAladdin = aladdinPlanAmt > eps;
        boolean hasPlan = planAmt > eps;
        if (hasActualColumn) {
            return formatStage3FixedSlotsAsText(
                    buildStage3QtyFixedLineSlots(
                            aladdinPlanAmt,
                            planAmt,
                            actualAmt,
                            stage3RevisedAfterTrial,
                            eps,
                            stage21CompareMode,
                            stage3BaselineActualAmt,
                            variant),
                    singleLineDisplay);
        }
        if (!hasAladdin && !hasPlan) {
            return "";
        }
        String aladdinFmt = hasAladdin ? ResultDispatchNormalizer.formatQty(aladdinPlanAmt) : "";
        if (hasPlan) {
            return formatDispatchPlanActualQtyText(
                    aladdinFmt,
                    ResultDispatchNormalizer.formatQty(planAmt),
                    "",
                    singleLineDisplay,
                    DispatchPlanQtyLineLabel.STAGE2);
        }
        return LABEL_ALADDIN_PLAN + aladdinFmt;
    }

    /** 配台数量行の括弧ラベル（段階2後／段階3前／なし）。 */
    enum DispatchPlanQtyLineLabel {
        PLAIN,
        STAGE2,
        STAGE3
    }

    static String formatDispatchPlanActualQtyText(
            String aladdinFmt, String planFmt, String actualFmt, boolean singleLineDisplay) {
        return formatDispatchPlanActualQtyText(
                aladdinFmt, planFmt, actualFmt, singleLineDisplay, DispatchPlanQtyLineLabel.STAGE3);
    }

    static String formatDispatchPlanActualQtyText(
            String aladdinFmt,
            String planFmt,
            String actualFmt,
            boolean singleLineDisplay,
            DispatchPlanQtyLineLabel planLabel) {
        StringBuilder sb = new StringBuilder();
        if (aladdinFmt != null && !aladdinFmt.isBlank()) {
            sb.append(LABEL_ALADDIN_PLAN).append(aladdinFmt);
        }
        if (planFmt != null && !planFmt.isBlank()) {
            if (!sb.isEmpty()) {
                sb.append(singleLineDisplay ? ' ' : '\n');
            }
            sb.append(planQtyLinePrefix(planLabel)).append(planFmt);
        }
        if (actualFmt != null && !actualFmt.isBlank()) {
            if (!sb.isEmpty()) {
                sb.append(singleLineDisplay ? ' ' : '\n');
            }
            sb.append(LABEL_STAGE3_ACTUAL).append(actualFmt);
        }
        return sb.toString();
    }

    private static String planQtyLinePrefix(DispatchPlanQtyLineLabel planLabel) {
        if (planLabel == null) {
            return "";
        }
        return switch (planLabel) {
            case STAGE2 -> LABEL_STAGE2_PLAN;
            case STAGE3 -> LABEL_STAGE3_PLAN;
            case PLAIN -> "";
        };
    }

    private static void applyDispatchStage3QtyLineLabelStyle(Label lbl, String line) {
        if (ResultDispatchStage3Support.isStage3RevisedQtyLine(line)) {
            lbl.getStyleClass().add(DISPATCH_STAGE3_REVISED_LINE_STYLE_CLASS);
            lbl.setStyle(STAGE3_REVISED_LINE_INLINE_STYLE);
        } else if (line.startsWith(LABEL_STAGE21_ACTUAL)) {
            lbl.getStyleClass().add(DISPATCH_STAGE21_AFTER_LINE_STYLE_CLASS);
            lbl.setStyle(STAGE21_AFTER_LINE_INLINE_STYLE);
        } else if (ResultDispatchStage3Support.isStage3AfterQtyLine(line)) {
            lbl.getStyleClass().add(DISPATCH_STAGE3_AFTER_LINE_STYLE_CLASS);
            lbl.setStyle(STAGE3_AFTER_LINE_INLINE_STYLE);
        } else {
            lbl.setStyle(STAGE3_QTY_DEFAULT_LINE_INLINE_STYLE);
        }
    }

    /**
     * 段階3ラベル Graphic 表示。item は空（納期管理 triple セルと同趣旨）し、コピー用テキストは property へ。
     */
    private static void applyDispatchQtyGraphicCellDisplay(
            SpreadsheetCell cell, Node graphic, String clipboardText) {
        cell.setCellGraphic(false);
        cell.setGraphic(graphic);
        if (!cell.getStyleClass().contains(DISPATCH_DATE_QTY_GRAPHIC_ONLY_STYLE_CLASS)) {
            cell.getStyleClass().add(DISPATCH_DATE_QTY_GRAPHIC_ONLY_STYLE_CLASS);
        }
        SpreadsheetTabularSupport.setSpreadsheetCellDisplayValue(cell, "");
        if (graphic != null) {
            Tooltip.install(graphic, new Tooltip(clipboardText != null ? clipboardText : ""));
        }
    }

    private static void clearDispatchQtyCellGraphic(SpreadsheetCell cell) {
        cell.setCellGraphic(false);
        cell.setGraphic(null);
        cell.getStyleClass().remove(DISPATCH_DATE_QTY_GRAPHIC_ONLY_STYLE_CLASS);
    }

    private static void tagDispatchDateQtyCell(SpreadsheetCell cell, boolean multiline) {
        tagDispatchDateQtyCell(cell, multiline, false);
    }

    private static void tagDispatchDateQtyCell(
            SpreadsheetCell cell, boolean multiline, boolean stage21QtyChanged) {
        if (cell == null) {
            return;
        }
        if (!cell.getStyleClass().contains(DISPATCH_DATE_QTY_CELL_STYLE_CLASS)) {
            cell.getStyleClass().add(DISPATCH_DATE_QTY_CELL_STYLE_CLASS);
        }
        cell.getStyleClass().remove(DISPATCH_DATE_QTY_SHORTFALL_CELL_STYLE_CLASS);
        if (stage21QtyChanged) {
            if (!cell.getStyleClass().contains(DISPATCH_STAGE21_QTY_CHANGED_CELL_STYLE_CLASS)) {
                cell.getStyleClass().add(DISPATCH_STAGE21_QTY_CHANGED_CELL_STYLE_CLASS);
            }
        } else {
            cell.getStyleClass().remove(DISPATCH_STAGE21_QTY_CHANGED_CELL_STYLE_CLASS);
        }
        setDispatchDateQtyMultilineStyleClass(cell, multiline);
    }

    private static void tagDispatchDateQtyShortfallCell(SpreadsheetCell cell, boolean multiline) {
        if (cell == null) {
            return;
        }
        if (!cell.getStyleClass().contains(DISPATCH_DATE_QTY_SHORTFALL_CELL_STYLE_CLASS)) {
            cell.getStyleClass().add(DISPATCH_DATE_QTY_SHORTFALL_CELL_STYLE_CLASS);
        }
        cell.getStyleClass().remove(DISPATCH_DATE_QTY_CELL_STYLE_CLASS);
        setDispatchDateQtyMultilineStyleClass(cell, multiline);
    }

    private static void setDispatchDateQtyMultilineStyleClass(SpreadsheetCell cell, boolean multiline) {
        if (multiline) {
            if (!cell.getStyleClass().contains(DISPATCH_DATE_QTY_MULTILINE_CELL_STYLE_CLASS)) {
                cell.getStyleClass().add(DISPATCH_DATE_QTY_MULTILINE_CELL_STYLE_CLASS);
            }
        } else {
            cell.getStyleClass().remove(DISPATCH_DATE_QTY_MULTILINE_CELL_STYLE_CLASS);
        }
    }

    private void installDispatchShortfallColumns(TableView<DispatchQtyShortfallRow> tv) {
        if (tv == null) {
            return;
        }
        tv.getColumns().clear();
        TableColumn<DispatchQtyShortfallRow, String> c0 = new TableColumn<>("依頼NO");
        c0.setCellValueFactory(
                cd -> new ReadOnlyObjectWrapper<>(Objects.toString(cd.getValue().taskId(), "")));
        TableColumn<DispatchQtyShortfallRow, String> c1 = new TableColumn<>("機械名");
        c1.setCellValueFactory(
                cd -> new ReadOnlyObjectWrapper<>(Objects.toString(cd.getValue().machineName(), "")));
        TableColumn<DispatchQtyShortfallRow, String> c2 = new TableColumn<>("配台日");
        c2.setCellValueFactory(
                cd ->
                        new ReadOnlyObjectWrapper<>(
                                Objects.toString(cd.getValue().dispatchDateIso(), "")));
        TableColumn<DispatchQtyShortfallRow, String> c3 = new TableColumn<>("目標(m)");
        c3.setCellValueFactory(
                cd ->
                        new ReadOnlyObjectWrapper<>(
                                formatShortfallMeters(cd.getValue().targetM())));
        TableColumn<DispatchQtyShortfallRow, String> c4 = new TableColumn<>("実績(m)");
        c4.setCellValueFactory(
                cd ->
                        new ReadOnlyObjectWrapper<>(
                                formatShortfallMeters(cd.getValue().doneM())));
        TableColumn<DispatchQtyShortfallRow, String> c5 = new TableColumn<>("不足(m)");
        c5.setCellValueFactory(
                cd ->
                        new ReadOnlyObjectWrapper<>(
                                formatShortfallMeters(cd.getValue().shortfallM())));
        TableColumn<DispatchQtyShortfallRow, String> c6 = new TableColumn<>("補足");
        c6.setPrefWidth(280);
        c6.setCellValueFactory(
                cd -> new ReadOnlyObjectWrapper<>(Objects.toString(cd.getValue().note(), "")));
        tv.getColumns().addAll(c0, c1, c2, c3, c4, c5, c6);
    }

    private static String formatShortfallMeters(double m) {
        if (Double.isNaN(m) || Double.isInfinite(m)) {
            return "";
        }
        if (Math.abs(m - Math.rint(m)) < 1e-6) {
            return Long.toString((long) Math.rint(m));
        }
        return String.format("%.3f", m);
    }

    /** {@code 結果_配台表.json} と同じフォルダの {@code dispatch_trial_shortages.json} から未達行を読み UI に反映する。 */
    private void applyDispatchShortfallFromDisk(Path resultDispatchJson) {
        List<DispatchQtyShortfallRow> rows = List.of();
        lastDispatchShortageHints = List.of();
        if (resultDispatchJson != null) {
            Path shortagePath = resultDispatchJson.resolveSibling("dispatch_trial_shortages.json");
            if (Files.isRegularFile(shortagePath)) {
                try {
                    lastShortagesNote = DispatchTrialShortages.readNote(shortagePath);
                    DispatchTrialShortages.FullBundle fb =
                            DispatchTrialShortages.readFull(shortagePath);
                    rows = fb.dispatchQtyShortfall();
                    lastDispatchShortageHints = List.copyOf(fb.shortageHints());
                } catch (IOException e) {
                    rows = List.of();
                    lastDispatchShortageHints = List.of();
                    lastShortagesNote = "";
                }
            } else {
                lastShortagesNote = "";
            }
        }
        rows = mergeDispatchQtyShortfallRowsUnique(
                rows, DispatchTimelineMetaMissShortfalls.detectFromDocument(doc));
        applyDispatchShortfallRows(rows);
        refreshStage21AttendanceApplyPanel(resultDispatchJson);
    }

    private static List<DispatchQtyShortfallRow> mergeDispatchQtyShortfallRowsUnique(
            List<DispatchQtyShortfallRow> primary, List<DispatchQtyShortfallRow> extra) {
        Map<String, DispatchQtyShortfallRow> byKey = new LinkedHashMap<>();
        if (primary != null) {
            for (DispatchQtyShortfallRow r : primary) {
                byKey.put(
                        DispatchTrialShortages.wideShortfallKey(
                                r.taskId(), r.machineName(), r.dispatchDateIso()),
                        r);
            }
        }
        if (extra != null) {
            for (DispatchQtyShortfallRow r : extra) {
                byKey.putIfAbsent(
                        DispatchTrialShortages.wideShortfallKey(
                                r.taskId(), r.machineName(), r.dispatchDateIso()),
                        r);
            }
        }
        return List.copyOf(byKey.values());
    }

    private void applyDispatchShortfallRows(List<DispatchQtyShortfallRow> rows) {
        lastDispatchShortfallRows = rows != null ? List.copyOf(rows) : List.of();
        dispatchWideShortfallKeys.clear();
        dispatchByDayShortfallKeys.clear();
        for (DispatchQtyShortfallRow r : lastDispatchShortfallRows) {
            dispatchWideShortfallKeys.add(
                    DispatchTrialShortages.wideShortfallKey(
                            r.taskId(), r.machineName(), r.dispatchDateIso()));
            dispatchByDayShortfallKeys.add(
                    DispatchTrialShortages.byDayShortfallKey(
                            r.machineName(), r.dispatchDateIso()));
        }
        if (dispatchShortfallTable != null) {
            dispatchShortfallTable.getItems().setAll(lastDispatchShortfallRows);
        }
        boolean vis = !lastDispatchShortfallRows.isEmpty();
        if (dispatchShortfallPanel != null) {
            dispatchShortfallPanel.setVisible(vis);
            dispatchShortfallPanel.setManaged(vis);
        }
    }

    private void clearDispatchShortfallUi() {
        lastDispatchShortfallRows = List.of();
        lastDispatchShortageHints = List.of();
        dispatchWideShortfallKeys.clear();
        dispatchByDayShortfallKeys.clear();
        if (dispatchShortfallTable != null) {
            dispatchShortfallTable.getItems().clear();
        }
        if (dispatchShortfallPanel != null) {
            dispatchShortfallPanel.setVisible(false);
            dispatchShortfallPanel.setManaged(false);
        }
    }

    private void showRawInputMorningDispatchRateWarningIfNeeded(Stage owner) {
        if (shell == null) {
            return;
        }
        PlanInputTabController planInput = shell.planInputTabControllerForDispatchRollUnit();
        if (planInput == null) {
            return;
        }
        var rawDates = planInput.collectEffectiveRawInputDateByTaskId();
        if (rawDates.isEmpty()) {
            return;
        }
        Path jsonPath = AppPaths.resolveResultDispatchTableJsonPath(shell.snapshotUiEnv());
        RawInputMorningDispatchRateWarning.showIfNeeded(shell, owner, jsonPath, rawDates);
    }

    /**
     * 配台試行（段階3）成功後、未達があるときにモーダル表示する。{@link DispatchTrialUnassignedWizard} より前。
     *
     * <p>手動確認の目安: 機械カレンダー等でブロックされる暦日に当日配台数量を置いて試行し、該当の日付セルが赤表示になり、
     * ツールバー下サマリ表と本ダイアログに同一内容が並ぶこと。
     */
    private void showDispatchQtyShortfallDialogIfNeeded(Stage owner) {
        if (lastDispatchShortfallRows == null || lastDispatchShortfallRows.isEmpty()) {
            return;
        }
        TableView<DispatchQtyShortfallRow> tv = new TableView<>();
        installDispatchShortfallColumns(tv);
        tv.getItems().setAll(lastDispatchShortfallRows);
        tv.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_FLEX_LAST_COLUMN);

        Label head =
                new Label(
                        "次の暦日で、手動修正表の配台目標（当日配台数量）に対しタイムライン実績が不足しています。"
                                + " 段階3では後日への後ろ倒し・カレンダー制約等で計画日未達になる場合があります。"
                                + " 該当セルは赤表示され、本一覧と同一内容です。");
        head.setWrapText(true);
        head.setStyle("-fx-font-size: 13px;");
        BorderPane root = new BorderPane();
        root.setTop(head);
        BorderPane.setMargin(head, new Insets(10, 14, 8, 14));
        root.setCenter(tv);
        BorderPane.setMargin(tv, new Insets(0, 14, 14, 14));

        Stage st = new Stage();
        if (owner != null) {
            st.initOwner(owner);
        }
        st.initModality(Modality.APPLICATION_MODAL);
        st.setTitle("配台数量未達（タイムライン実績）");
        Scene sc = new Scene(root, 920, 520);
        if (shell != null) {
            shell.registerThemeTrackedScene(sc);
        }
        st.setScene(sc);
        st.setOnHidden(
                ev -> {
                    if (shell != null) {
                        shell.unregisterThemeTrackedScene(sc);
                    }
                });
        st.showAndWait();
    }

    /**
     * 配台試行後、{@code dispatch_trial_shortages.json} の op_shortage / as_shortage が空でなければモーダルで示す。
     * メートル未達（{@link #showDispatchQtyShortfallDialogIfNeeded}）とは別系統。
     */
    private void showDispatchShortageHintsDialogIfNeeded(Stage owner) {
        if (lastDispatchShortageHints == null || lastDispatchShortageHints.isEmpty()) {
            return;
        }
        TableView<DispatchTrialShortages.ShortageHint> tv = new TableView<>();
        tv.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_FLEX_LAST_COLUMN);

        TableColumn<DispatchTrialShortages.ShortageHint, String> h0 = new TableColumn<>("依頼NO");
        h0.setCellValueFactory(
                cd ->
                        new ReadOnlyObjectWrapper<>(
                                Objects.toString(cd.getValue().taskId(), "")));
        TableColumn<DispatchTrialShortages.ShortageHint, String> h1 = new TableColumn<>("理由");
        h1.setPrefWidth(220);
        h1.setCellValueFactory(
                cd ->
                        new ReadOnlyObjectWrapper<>(
                                Objects.toString(cd.getValue().reason(), "")));
        TableColumn<DispatchTrialShortages.ShortageHint, String> h2 = new TableColumn<>("補足");
        h2.setPrefWidth(380);
        h2.setCellValueFactory(
                cd ->
                        new ReadOnlyObjectWrapper<>(
                                Objects.toString(cd.getValue().detail(), "")));
        tv.getColumns().addAll(h0, h1, h2);
        tv.getItems().setAll(lastDispatchShortageHints);

        Label head =
                new Label(
                        "フォーム候補不足（op_shortage）または割当不可（as_shortage）として記録された件です。"
                                + " 段階3（段階2同一）の試行ではこの種別は出ません（後ろ倒しで解消する前提）。"
                                + " 配台できないのは機械カレンダー・勤怠が master に無い場合のみです。");
        head.setWrapText(true);
        head.setStyle("-fx-font-size: 13px;");
        BorderPane root = new BorderPane();
        root.setTop(head);
        BorderPane.setMargin(head, new Insets(10, 14, 8, 14));
        root.setCenter(tv);
        BorderPane.setMargin(tv, new Insets(0, 14, 14, 14));

        Stage st = new Stage();
        if (owner != null) {
            st.initOwner(owner);
        }
        st.initModality(Modality.APPLICATION_MODAL);
        st.setTitle("人員・割当不足（試行スナップショット）");
        Scene sc = new Scene(root, 920, 480);
        if (shell != null) {
            shell.registerThemeTrackedScene(sc);
        }
        st.setScene(sc);
        st.setOnHidden(
                ev -> {
                    if (shell != null) {
                        shell.unregisterThemeTrackedScene(sc);
                    }
                });
        st.showAndWait();
    }

    /** タスク×日付: 日付列ダブルクリックで {@link #LABEL_STAGE3_PLAN} のみ編集（{@link #LABEL_STAGE3_ACTUAL} は固定）。 */
    private void installWideDoubleClickHandler() {
        wideSpreadsheet.addEventFilter(
                MouseEvent.MOUSE_CLICKED,
                e -> {
                    if (e.getClickCount() != 2) {
                        return;
                    }
                    TableCell<?, ?> tc = findTableCell(e.getPickResult().getIntersectedNode());
                    if (tc == null || !isUnderSpreadsheet(wideSpreadsheet, tc)) {
                        return;
                    }
                    int modelCol = wideModelColumnFromTableCell(tc);
                    int staticCols = WIDE_STATIC_HEADERS.size();
                    if (modelCol < staticCols) {
                        return;
                    }
                    int slot = modelCol - staticCols;
                    int dateIdx = slot / DAY_SLOT_COLUMNS;
                    int profIdx = wideProfileIndexFromTableCell(tc);
                    if (profIdx < 0
                            || profIdx >= wideRowItems.size()
                            || dateIdx < 0
                            || dateIdx >= dateAxis.size()) {
                        return;
                    }
                    WideRow wr = wideRowItems.get(profIdx);
                    double planQ = wr.getAmount(dateIdx);
                    double actualQ = wr.getActualAmount(dateIdx);
                    boolean hasActual = docHasActualDispatchQtyColumn() && actualQ > 1e-9;
                    TextInputDialog dialog =
                            new TextInputDialog(
                                    planQ > 1e-9 ? ResultDispatchNormalizer.formatQty(planQ) : "");
                    if (shell != null) {
                        dialog.initOwner(shell.primaryStageForDialogs());
                    }
                    dialog.setTitle("当日配台数量（段階3前）");
                    String profileHint = rollUnitProfileHint(wr, dateAxis.get(dateIdx));
                    Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM unitInfo =
                            resolveRollUnitForWideRow(wr);
                    StringBuilder header = new StringBuilder(profileHint);
                    if (hasActual) {
                        String actualLabel =
                                effectiveStage3PlanningVariant().actualQtyLabel();
                        header.append('\n')
                                .append(actualLabel)
                                .append(ResultDispatchNormalizer.formatQty(actualQ))
                                .append("（固定・ドラッグ移動不可）");
                    }
                    header.append('\n')
                            .append(
                                    DispatchInteractiveRollUnitSupport.rollUnitDialogHeader(
                                            Math.max(planQ, 0.0), unitInfo, null));
                    dialog.setHeaderText(header.toString());
                    dialog.setContentText(
                            (hasActual ? LABEL_STAGE3_PLAN : "数量")
                                    + " (m) — 配台ロール単位の整数倍のみ:");
                    Optional<String> ov = dialog.showAndWait();
                    ov.filter(s -> !s.isBlank())
                            .flatMap(
                                    s ->
                                            DispatchInteractiveRollUnitSupport
                                                    .parseRollAlignedTotalQuantity(
                                                            shell != null
                                                                    ? shell.primaryStageForDialogs()
                                                                    : null,
                                                            s,
                                                            unitInfo,
                                                            rollUnitProfileHint(
                                                                    wr, dateAxis.get(dateIdx))))
                            .ifPresent(
                                    newTotal -> {
                                        ResultDispatchPivot.upsertAllocationForWideMerge(
                                                doc.columns(),
                                                doc.rows(),
                                                wr.profileMap(),
                                                dateAxis.get(dateIdx),
                                                newTotal,
                                                ResultDispatchPivot
                                                        .DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
                                        ResultDispatchNormalizer.normalizeInPlace(
                                                doc.columns(), doc.rows());
                                        rebuildGrids();
                                        markDispatchDocDirty();
                                    });
                });
    }

    private void installWideDnDHandlers() {
        wideSpreadsheet.addEventFilter(
                MouseEvent.DRAG_DETECTED,
                e -> {
                    TableCell<?, ?> tc = findTableCell(e.getPickResult().getIntersectedNode());
                    if (tc == null || !isUnderSpreadsheet(wideSpreadsheet, tc)) {
                        return;
                    }
                    int modelCol = wideModelColumnFromTableCell(tc);
                    int staticCols = WIDE_STATIC_HEADERS.size();
                    if (modelCol < 0) {
                        return;
                    }

                    // Row reorder: drag from leading (static) model columns ? start gesture from the TableCell node.
                    if (modelCol < staticCols) {
                        int profIdx = wideProfileIndexFromTableCell(tc);
                        if (profIdx < 0 || profIdx >= wideProfiles.size()) {
                            return;
                        }
                        List<String> cols = doc.columns();
                        String gk = ResultDispatchNormalizer.staticGroupKey(cols, wideProfiles.get(profIdx));
                        String b64 =
                                Base64.getUrlEncoder()
                                        .withoutPadding()
                                        .encodeToString(gk.getBytes(StandardCharsets.UTF_8));
                        Dragboard db = tc.startDragAndDrop(TransferMode.MOVE);
                        ClipboardContent cc = new ClipboardContent();
                        cc.putString(DND_ROW_PREFIX + b64);
                        db.setContent(cc);
                        SpreadsheetRowReorderDragGhost.apply(db, tc, e);
                        e.consume();
                        return;
                    }

                    int slot = modelCol - staticCols;
                    int dateIdx = slot / DAY_SLOT_COLUMNS;
                    int profIdx = wideProfileIndexFromTableCell(tc);
                    if (profIdx < 0 || profIdx >= wideRowItems.size() || dateIdx < 0 || dateIdx >= dateAxis.size()) {
                        return;
                    }
                    WideRow wr = wideRowItems.get(profIdx);
                    double qty = wr.getAmount(dateIdx);
                    if (qty <= 1e-9) {
                        return;
                    }
                    Dragboard db = tc.startDragAndDrop(TransferMode.MOVE);
                    ClipboardContent cc = new ClipboardContent();
                    List<String> cols = doc.columns();
                    String gk = ResultDispatchNormalizer.staticGroupKey(cols, wr.profileMap());
                    String b64 =
                            Base64.getUrlEncoder()
                                    .withoutPadding()
                                    .encodeToString(gk.getBytes(StandardCharsets.UTF_8));
                    cc.putString(DND_PREFIX + DND_V2_MARKER + b64 + ":" + dateIdx + ":" + qty);
                    db.setContent(cc);
                    e.consume();
                });

        wideSpreadsheet.addEventFilter(
                DragEvent.DRAG_OVER,
                e -> {
                    TableCell<?, ?> tc = findTableCell(e.getPickResult().getIntersectedNode());
                    if (tc == null || !isUnderSpreadsheet(wideSpreadsheet, tc)) {
                        return;
                    }
                    int modelCol = wideModelColumnFromTableCell(tc);
                    int staticCols = WIDE_STATIC_HEADERS.size();
                    if (modelCol < 0) {
                        return;
                    }

                    if (modelCol < staticCols) {
                        if (e.getDragboard().hasString()
                                && e.getDragboard().getString().startsWith(DND_ROW_PREFIX)) {
                            int profIdx = wideProfileIndexFromTableCell(tc);
                            if (profIdx >= 0 && profIdx < wideProfiles.size()) {
                                e.acceptTransferModes(TransferMode.MOVE);
                            }
                        }
                        e.consume();
                        return;
                    }

                    int profIdx = wideProfileIndexFromTableCell(tc);
                    if (profIdx < 0 || profIdx >= wideRowItems.size()) {
                        return;
                    }
                    if (e.getDragboard().hasString()
                            && e.getDragboard().getString().startsWith(DND_PREFIX)) {
                        e.acceptTransferModes(TransferMode.MOVE);
                    }
                    e.consume();
                });

        wideSpreadsheet.addEventFilter(
                DragEvent.DRAG_DROPPED,
                e -> {
                    TableCell<?, ?> tc = findTableCell(e.getPickResult().getIntersectedNode());
                    if (tc == null || !isUnderSpreadsheet(wideSpreadsheet, tc)) {
                        return;
                    }
                    int modelCol = wideModelColumnFromTableCell(tc);
                    int staticCols = WIDE_STATIC_HEADERS.size();
                    if (modelCol < 0) {
                        return;
                    }

                    if (modelCol < staticCols) {
                        String payload = e.getDragboard().getString();
                        if (payload != null && payload.startsWith(DND_ROW_PREFIX)) {
                            boolean ok = handleWideRowReorderDrop(payload, tc);
                            e.setDropCompleted(ok);
                        } else {
                            e.setDropCompleted(false);
                        }
                        e.consume();
                        return;
                    }

                    int slot = modelCol - staticCols;
                    int dateIdx = slot / DAY_SLOT_COLUMNS;
                    int profIdx = wideProfileIndexFromTableCell(tc);
                    if (profIdx < 0 || profIdx >= wideRowItems.size()) {
                        return;
                    }
                    WideRow tgt = wideRowItems.get(profIdx);
                    boolean ok = handleWideDrop(e.getDragboard().getString(), tgt, dateIdx);
                    e.setDropCompleted(ok);
                    e.consume();
                });
    }

    private void installByDayDoubleClickHandler() {
        byDaySpreadsheet.addEventFilter(
                MouseEvent.MOUSE_CLICKED,
                e -> {
                    if (e.getClickCount() != 2) {
                        return;
                    }
                    TableCell<?, ?> tc = findTableCell(e.getPickResult().getIntersectedNode());
                    if (tc == null || !isUnderSpreadsheet(byDaySpreadsheet, tc)) {
                        return;
                    }
                    int col = tc.getTableView().getColumns().indexOf(tc.getTableColumn());
                    int staticCols = BY_DAY_STATIC_HEADERS.size();
                    if (col < staticCols) {
                        return;
                    }
                    int slot = col - staticCols;
                    int dateIdx = slot / DAY_SLOT_COLUMNS;
                    int row = tc.getIndex();
                    int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
                    int dataIdx = row - firstData;
                    List<Map.Entry<String, String>> keys = ResultDispatchPivot.sortedProcessMachineKeys(doc.rows());
                    if (dataIdx < 0 || dataIdx >= keys.size()) {
                        return;
                    }
                    Map.Entry<String, String> en = keys.get(dataIdx);
                    double cur =
                            ResultDispatchPivot.sumQuantityForProcessMachineDate(
                                    doc.rows(), en.getKey(), en.getValue(), dateAxis.get(dateIdx));
                    TextInputDialog dialog =
                            new TextInputDialog(ResultDispatchNormalizer.formatQty(cur));
                    dialog.initOwner(shell != null ? shell.getPrimaryStage() : null);
                    dialog.setTitle("日別合計");
                    dialog.setHeaderText(
                            en.getKey()
                                    + " / "
                                    + en.getValue()
                                    + " / "
                                    + dateAxis.get(dateIdx));
                    Optional<String> ov = dialog.showAndWait();
                    ov.filter(s -> !s.isBlank())
                            .ifPresent(
                                    s -> {
                                        double newTotal = ResultDispatchNormalizer.parseDouble(s);
                                        ResultDispatchPivot.scaleProcessMachineDateToTotal(
                                                doc.columns(),
                                                doc.rows(),
                                                en.getKey(),
                                                en.getValue(),
                                                dateAxis.get(dateIdx),
                                                newTotal);
                                        ResultDispatchNormalizer.normalizeInPlace(doc.columns(), doc.rows());
                                        rebuildGrids();
                                        markDispatchDocDirty();
                                    });
                });
    }

    private static TableCell<?, ?> findTableCell(Node n) {
        while (n != null) {
            if (n instanceof TableCell<?, ?> tc) {
                return tc;
            }
            n = n.getParent();
        }
        return null;
    }

    private static boolean isUnderSpreadsheet(SpreadsheetView spv, Node node) {
        Node n = node;
        while (n != null) {
            if (n == spv) {
                return true;
            }
            n = n.getParent();
        }
        return false;
    }

    private int wideProfileIndexForRow(WideRow row) {
        List<String> cols = doc.columns();
        String gk = ResultDispatchNormalizer.staticGroupKey(cols, row.profileMap());
        return indexOfProfileGroupKey(gk);
    }

    private int indexOfProfileGroupKey(String groupKey) {
        List<String> cols = doc.columns();
        for (int i = 0; i < wideProfiles.size(); i++) {
            if (ResultDispatchNormalizer.staticGroupKey(cols, wideProfiles.get(i)).equals(groupKey)) {
                return i;
            }
        }
        return -1;
    }

    private boolean handleWideRowReorderDrop(String payload, TableCell<?, ?> targetCell) {
        if (payload == null || !payload.startsWith(DND_ROW_PREFIX)) {
            return false;
        }
        String b64 = payload.substring(DND_ROW_PREFIX.length());
        final String gk;
        try {
            gk = new String(Base64.getUrlDecoder().decode(b64), StandardCharsets.UTF_8);
        } catch (IllegalArgumentException ex) {
            return false;
        }
        int fromIdx = indexOfProfileGroupKey(gk);
        int toIdx = wideProfileIndexFromTableCell(targetCell);
        if (fromIdx < 0 || toIdx < 0 || fromIdx >= wideProfiles.size() || toIdx >= wideProfiles.size()) {
            return false;
        }
        if (fromIdx == toIdx) {
            return false;
        }
        wideProfiles.add(toIdx, wideProfiles.remove(fromIdx));
        assignSequentialTrialOrdersForProfiles(wideProfiles);
        ResultDispatchNormalizer.normalizeInPlace(doc.columns(), doc.rows());
        rebuildGrids();
        statusLabel.setText("行を移動しました");
        markDispatchDocDirty();
        return true;
    }

    private boolean handleWideDrop(String payload, WideRow targetRow, int targetDateIdx) {
        if (!payload.startsWith(DND_PREFIX)) {
            return false;
        }
        String rest = payload.substring(DND_PREFIX.length());
        boolean payloadIsV2 = rest.startsWith(DND_V2_MARKER);
        List<String> cols = doc.columns();

        int fromRow;
        int fromDateIdx;
        double max;

        if (payloadIsV2) {
            String body = rest.substring(DND_V2_MARKER.length());
            String[] p = body.split(":", 3);
            if (p.length < 3) {
                return false;
            }
            try {
                String gk =
                        new String(Base64.getUrlDecoder().decode(p[0]), StandardCharsets.UTF_8);
                fromDateIdx = Integer.parseInt(p[1]);
                max = Double.parseDouble(p[2]);
                fromRow = indexOfProfileGroupKey(gk);
            } catch (IllegalArgumentException e) {
                return false;
            }
            if (fromRow < 0) {
                return false;
            }
        } else {
            String[] p = rest.split(":");
            if (p.length < 3) {
                return false;
            }
            try {
                fromRow = Integer.parseInt(p[0]);
                fromDateIdx = Integer.parseInt(p[1]);
                max = Double.parseDouble(p[2]);
            } catch (NumberFormatException e) {
                return false;
            }
        }

        int toIdx = wideProfileIndexForRow(targetRow);

        if (fromRow != toIdx) {
            statusLabel.setText(
                    "縦方向への移動はできません（横のみ）");
            return false;
        }
        if (fromRow == toIdx && fromDateIdx == targetDateIdx) {
            return false;
        }
        if (fromRow < 0
                || fromRow >= wideProfiles.size()
                || fromRow >= wideRowItems.size()
                || toIdx < 0
                || toIdx >= wideProfiles.size()) {
            return false;
        }
        WideRow fromWr = wideRowItems.get(fromRow);
        Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM unitInfo =
                resolveRollUnitForWideRow(fromWr);
        String profileHint = rollUnitProfileHint(fromWr, dateAxis.get(fromDateIdx));
        Optional<Double> moved =
                DispatchInteractiveRollUnitSupport.pickRollAlignedMoveQuantity(
                        shell != null ? shell.primaryStageForDialogs() : null,
                        max,
                        unitInfo,
                        profileHint);
        if (moved.isEmpty()) {
            return false;
        }
        double amt = moved.get();
        if (amt <= 1e-9
                || amt > max + 1e-9
                || !Stage2PlanRowDispatchQtyMetrics.isQtyAlignedToRollUnit(
                        amt, unitInfo.unitM())) {
            return false;
        }
        Map<String, String> fromProfile = wideProfiles.get(fromRow);
        Map<String, String> toProfile = wideProfiles.get(toIdx);
        LocalDate fromDay = dateAxis.get(fromDateIdx);
        LocalDate toDay = dateAxis.get(targetDateIdx);

        double fromSum =
                ResultDispatchPivot.sumQuantityForProfileAndDateForWideMerge(
                        doc.rows(),
                        fromProfile,
                        fromDay,
                        ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
        double toSum =
                ResultDispatchPivot.sumQuantityForProfileAndDateForWideMerge(
                        doc.rows(),
                        toProfile,
                        toDay,
                        ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
        ResultDispatchPivot.upsertAllocationForWideMerge(
                cols,
                doc.rows(),
                fromProfile,
                fromDay,
                fromSum - amt,
                ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
        ResultDispatchPivot.upsertAllocationForWideMerge(
                cols,
                doc.rows(),
                toProfile,
                toDay,
                toSum + amt,
                ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
        ResultDispatchNormalizer.normalizeInPlace(cols, doc.rows());
        statusLabel.setText(
                "移動: "
                        + ResultDispatchNormalizer.formatQty(amt)
                        + " m（ロール単位 "
                        + ResultDispatchNormalizer.formatQty(unitInfo.unitM())
                        + " m）");
        rebuildGrids();
        markDispatchDocDirty();
        return true;
    }

    private Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM resolveRollUnitForWideRow(
            WideRow wr) {
        PlanInputTabController planInput =
                shell != null ? shell.planInputTabControllerForDispatchRollUnit() : null;
        Map<String, String> ui = shell != null ? shell.snapshotUiEnv() : Map.of();
        return DispatchInteractiveRollUnitSupport.resolveUnitM(
                wr.profileMap(), planInput, ui, rollUnitTablesCached());
    }

    private Stage2RollUnitLengthTables rollUnitTablesCached() {
        return cachedRollUnitTables.updateAndGet(
                cur -> {
                    if (cur != null) {
                        return cur;
                    }
                    if (shell == null) {
                        return Stage2RollUnitLengthTables.empty();
                    }
                    try {
                        return Stage2RollUnitLengthTables.load(shell.snapshotUiEnv());
                    } catch (Exception ignored) {
                        return Stage2RollUnitLengthTables.empty();
                    }
                });
    }

    private static String rollUnitProfileHint(WideRow wr, LocalDate day) {
        return Objects.toString(wr.getStatic("依頼NO"), "")
                + " / "
                + Objects.toString(wr.getStatic(ResultDispatchSchema.COL_MACHINE), "")
                + " / "
                + day;
    }

    /** Mutable wide row (amounts indexed by {@link #dateAxis}). */
    public static final class WideRow {
        private final Map<String, String> staticPart;
        private final double[] amounts;
        private final double[] actualAmounts;

        WideRow(Map<String, String> staticPart, int nDates) {
            this.staticPart = new LinkedHashMap<>(staticPart);
            this.amounts = new double[nDates];
            this.actualAmounts = new double[nDates];
        }

        String getStatic(String col) {
            return staticPart.getOrDefault(col, "");
        }

        Map<String, String> profileMap() {
            return new LinkedHashMap<>(staticPart);
        }

        double getAmount(int di) {
            return amounts[di];
        }

        void setAmount(int di, double v) {
            amounts[di] = v;
        }

        double getActualAmount(int di) {
            return actualAmounts[di];
        }

        void setActualAmount(int di, double v) {
            actualAmounts[di] = v;
        }

        double sumPlanAmounts() {
            double sum = 0;
            for (double v : amounts) {
                sum += v;
            }
            return sum;
        }

        double sumActualAmounts() {
            double sum = 0;
            for (double v : actualAmounts) {
                sum += v;
            }
            return sum;
        }
    }

    public record ByDayRow(
            String process, String machine, String processingContent, double[] amounts, double[] actualAmounts) {
        ByDayRow(String process, String machine, String processingContent, int n) {
            this(process, machine, processingContent, new double[n], new double[n]);
        }

        double getAmount(int i) {
            return amounts[i];
        }

        void setAmount(int i, double v) {
            amounts[i] = v;
        }

        double getActualAmount(int i) {
            return actualAmounts[i];
        }

        void setActualAmount(int i, double v) {
            actualAmounts[i] = v;
        }

        double sumPlanAmounts() {
            double sum = 0;
            for (double v : amounts) {
                sum += v;
            }
            return sum;
        }

        double sumActualAmounts() {
            double sum = 0;
            for (double v : actualAmounts) {
                sum += v;
            }
            return sum;
        }
    }
}
