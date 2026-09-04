package jp.co.pm.ai.desktop;

import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.format.TextStyle;
import java.time.temporal.ChronoUnit;
import java.time.temporal.TemporalAdjusters;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.atomic.AtomicLong;

import javafx.animation.KeyFrame;
import javafx.animation.PauseTransition;
import javafx.animation.Timeline;
import javafx.application.Platform;
import javafx.beans.property.ReadOnlyObjectWrapper;
import javafx.beans.property.SimpleBooleanProperty;
import javafx.collections.FXCollections;
import javafx.concurrent.Task;
import javafx.css.PseudoClass;
import javafx.fxml.FXML;
import javafx.geometry.Insets;
import javafx.geometry.Point2D;
import javafx.geometry.Pos;
import javafx.geometry.Side;
import javafx.scene.Node;
import javafx.scene.chart.BarChart;
import javafx.scene.chart.CategoryAxis;
import javafx.scene.chart.LineChart;
import javafx.scene.chart.XYChart;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ComboBox;
import javafx.scene.control.DatePicker;
import javafx.scene.control.Label;
import javafx.scene.control.ListCell;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableRow;
import javafx.scene.control.TableView;
import javafx.scene.control.ToggleButton;
import javafx.scene.control.ToggleGroup;
import javafx.scene.control.Tooltip;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyEvent;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Pane;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.scene.shape.Line;
import javafx.util.Duration;
import javafx.util.StringConverter;

import jp.co.pm.ai.desktop.ProcessingTrendChartSupport.NiceRange;
import jp.co.pm.ai.desktop.io.actuals.DashboardLoadErrorFormatter;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardSourceLoader;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardSourceLoader.LoadedSources;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardSourceLoader.ReloadDecision;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardSourceLoader.SourceFingerprint;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.DayPoint;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.Filter;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.PlanSource;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.Result;

/**
 * 「加工トレンド」タブ: 加工実績と加工予定を日別に重ねた複合グラフ（日次棒 + 累計折れ線）。
 *
 * <p>ソース読込はダッシュボードと同じ {@link EquipmentStatusDashboardSourceLoader} を使い、
 * 集計は {@link ProcessingTrendAggregator} に委ねる。読込・集計はいずれもバックグラウンドで行い、
 * FX スレッドでは描画のみ行う。常時表示を想定し、自動更新（指紋比較のみなので安価）と
 * 日跨ぎでの「今日」追従を持つ。
 */
public class ProcessingTrendTabController {

    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("yyyy/MM/dd");
    private static final DateTimeFormatter TIME_FMT = DateTimeFormatter.ofPattern("HH:mm:ss");
    private static final DateTimeFormatter MONTH_FMT = DateTimeFormatter.ofPattern("yyyy年M月");
    private static final String ALL_ITEM = "（すべて）";
    private static final double CHART_TOP_PADDING = 6.0;
    /** 自前 X ラベルの上端（軸ローカル y）: 目盛線の長さ 5px + 余白。 */
    private static final double X_LABEL_TOP_OFFSET = 7.0;
    /** UI で許す最長期間（日）。集計側の上限（1000 日）より手前で止め、棒が潰れる前に知らせる。 */
    static final int MAX_PERIOD_DAYS = 400;
    static final int AUTO_REFRESH_DEFAULT_SEC = 300;
    static final int AUTO_REFRESH_MIN_SEC = 30;
    static final int AUTO_REFRESH_MAX_SEC = 3600;
    /** 進捗率の色分け閾値（%）。 */
    static final double PROGRESS_GOOD_PCT = 100.0;
    static final double PROGRESS_WARN_PCT = 90.0;
    private static final PseudoClass PC_GOOD = PseudoClass.getPseudoClass("good");
    private static final PseudoClass PC_WARN = PseudoClass.getPseudoClass("warn");
    private static final PseudoClass PC_BAD = PseudoClass.getPseudoClass("bad");
    private static final PseudoClass PC_WEEKEND = PseudoClass.getPseudoClass("weekend");
    private static final PseudoClass PC_TODAY = PseudoClass.getPseudoClass("today");
    private static final PseudoClass PC_FUTURE = PseudoClass.getPseudoClass("future");
    private static final PseudoClass PC_SAT = PseudoClass.getPseudoClass("sat");
    private static final PseudoClass PC_SUN = PseudoClass.getPseudoClass("sun");

    enum PeriodPreset {
        THIS_MONTH("今月"),
        LAST_MONTH("先月"),
        NEXT_MONTH("来月"),
        AROUND_2_WEEKS("前後2週"),
        PAST_4_WEEKS("過去4週"),
        CUSTOM("任意");

        private final String label;

        PeriodPreset(String label) {
            this.label = label;
        }

        @Override
        public String toString() {
            return label;
        }
    }

    enum ViewMode {
        COMBO,
        DAILY,
        CUMULATIVE
    }

    /** バナーの種類。描画成功時にどれを消すかの判断に使う。 */
    private enum NoticeKind {
        NONE,
        LOAD_ERROR,
        COMPUTE_ERROR,
        /** 集計結果に付随するデータ上の注意（列欠落・実績ソースの期間不足）。 */
        DATA_WARN,
        INFO
    }

    @FXML private BorderPane tabRoot;
    @FXML private Button reloadButton;
    @FXML private ComboBox<PeriodPreset> periodPresetCombo;
    @FXML private DatePicker fromDatePicker;
    @FXML private DatePicker toDatePicker;
    @FXML private CheckBox autoRefreshCheckBox;
    @FXML private Spinner<Integer> autoRefreshIntervalSpinner;
    @FXML private Label nextRefreshLabel;
    @FXML private ComboBox<PlanSource> planSourceCombo;
    @FXML private ComboBox<String> machineCombo;
    @FXML private ComboBox<String> processCombo;
    @FXML private ToggleGroup viewModeGroup;
    @FXML private ToggleButton viewComboToggle;
    @FXML private ToggleButton viewDailyToggle;
    @FXML private ToggleButton viewCumulativeToggle;
    @FXML private ProgressIndicator loadingIndicator;
    @FXML private Label loadingStatusLabel;
    @FXML private Label lastUpdatedLabel;
    @FXML private Label sourceSummaryLabel;
    @FXML private HBox noticeBanner;
    @FXML private Label noticeBannerLabel;
    @FXML private Button noticeReloadButton;
    @FXML private Label kpiActualValue;
    @FXML private Label kpiActualSub;
    @FXML private Label kpiPlanValue;
    @FXML private Label kpiPlanSub;
    @FXML private Label kpiProgressValue;
    @FXML private Label kpiProgressSub;
    @FXML private Label kpiRemainingValue;
    @FXML private Label kpiRemainingSub;
    @FXML private VBox kpiProjectedCard;
    @FXML private Label kpiProjectedValue;
    @FXML private Label kpiProjectedSub;
    @FXML private Label chartTitleLabel;
    @FXML private HBox legendBox;
    @FXML private Label unitHintLabel;
    @FXML private BarChart<String, Number> dailyChart;
    @FXML private CategoryAxis dailyXAxis;
    @FXML private ProcessingTrendNumberAxis dailyYAxis;
    @FXML private LineChart<String, Number> cumulativeChart;
    @FXML private CategoryAxis cumulativeXAxis;
    @FXML private ProcessingTrendNumberAxis cumulativeYAxis;
    @FXML private Pane markerPane;
    @FXML private VBox emptyStatePane;
    @FXML private Label emptyStateTitle;
    @FXML private Label emptyStateDetail;
    @FXML private TableView<DayPoint> detailTable;
    @FXML private TableColumn<DayPoint, DayPoint> colDate;
    @FXML private TableColumn<DayPoint, Number> colPlan;
    @FXML private TableColumn<DayPoint, Number> colActual;
    @FXML private TableColumn<DayPoint, Number> colDiff;
    @FXML private TableColumn<DayPoint, Number> colPlanCum;
    @FXML private TableColumn<DayPoint, Number> colActualCum;
    @FXML private TableColumn<DayPoint, Number> colProjectedCum;

    private final ExecutorService pool =
            Executors.newSingleThreadExecutor(
                    r -> {
                        Thread t = new Thread(r, "processing-trend-reload");
                        t.setDaemon(true);
                        return t;
                    });

    private final XYChart.Series<String, Number> actualDailySeries = new XYChart.Series<>();
    private final XYChart.Series<String, Number> planDailySeries = new XYChart.Series<>();
    private final XYChart.Series<String, Number> actualCumSeries = new XYChart.Series<>();
    private final XYChart.Series<String, Number> planCumSeries = new XYChart.Series<>();
    private final XYChart.Series<String, Number> projectedCumSeries = new XYChart.Series<>();
    private final Line todayLine = new Line();
    private final Label todayMarkerLabel = new Label("今日");
    /**
     * X 軸ラベルは CategoryAxis に描かせず（final クラスで間引き不可・長期間で 90° 自動回転する）
     * オーバーレイに自前で置く。ノードは再利用する。
     */
    private final List<Label> xAxisLabelPool = new ArrayList<>();
    /** X 軸ラベルを出すカテゴリ index（{@link ProcessingTrendChartSupport#labelledDates}）。 */
    private Set<Integer> labelledCategoryIdx = Set.of();
    private List<LocalDate> currentDates = List.of();
    private final SimpleBooleanProperty loading = new SimpleBooleanProperty(false);
    private final AtomicLong computeGeneration = new AtomicLong();
    /** 期間◀▶連打・フィルタ連続変更で集計がキューに溜まらないようにする。 */
    private final PauseTransition recomputeDebounce = new PauseTransition(Duration.millis(150));
    /** データ点ごとに Tooltip を作らず 1 個を共有し、ホバー時に文言を差し替える。 */
    private final Tooltip sharedChartTooltip = new Tooltip();

    private MainShellController shell;
    /**
     * 読込 Task の多重起動ガード。{@code Task.isRunning()} は {@code pool.execute} 直後はまだ {@code READY} で、
     * 同一 FX パルス内の 2 回目の呼び出しを止められないため boolean で持つ。
     */
    private boolean reloadInFlight;
    /** 読込中・集計中は別フラグで持ち、{@code loading} はその OR（集計完了で読込中表示が消えないように）。 */
    private boolean reloadingFlag;
    private boolean computeInFlight;
    private Task<Result> activeComputeTask;
    private LoadedSources cachedSources;
    private SourceFingerprint loadedFingerprint;
    private Result currentResult;
    private List<String> currentCategoryLabels = List.of();
    private String lastLoadErrorDetail = "";
    private NoticeKind noticeKind = NoticeKind.NONE;
    private LocalTime lastSuccessAt;
    private LocalTime lastDataChangedAt;
    private LocalTime lastCheckedAt;
    private boolean suppressFilterEvents;
    private boolean markerLayoutScheduled;
    private boolean tabActive;
    private Timeline autoRefreshTimeline;
    private int autoRefreshRemainingSec;

    @FXML
    private void initialize() {
        initPeriodControls();
        initAutoRefreshControls();
        initFilterCombos();
        initViewModeToggles();
        initCharts();
        initDetailTable();
        initLegend();
        initKeyboard();
        recomputeDebounce.setOnFinished(e -> recomputeNow());
        sharedChartTooltip.setShowDelay(Duration.millis(120));
        sharedChartTooltip.setHideDelay(Duration.millis(80));
        loading.addListener((obs, o, n) -> applyLoadingPresentation(n));
        applyLoadingPresentation(false);
        applyPreset(PeriodPreset.THIS_MONTH);
        applyViewMode(ViewMode.COMBO);
        Tooltip.install(
                unitHintLabel,
                new Tooltip(
                        "実績・予定ともに換算数量（m）を依頼×工程×機械ごとに合算した「工程延べ m」。\n"
                                + "全工程を合算した値は依頼の生産量ではありません（同じ依頼が工程ごとに数えられます）。\n"
                                + "参考 Excel（売上金額・依頼単位・完了日基準）とは単位・件数基準が異なります。\n"
                                + "見込累計 = 前日までの実績累計 + 当日以降の予定（当日は実績と予定の大きい方）"));
        Tooltip.install(
                kpiProjectedCard,
                new Tooltip(
                        "見込合計 = 前日までの実績 + 当日以降の予定（当日のみ実績と予定の大きい方）。\n"
                                + "アラジン予定は当日以降を行ごとに「未加工」で上限し、完了行は 0 とします。\n"
                                + "下段は予定合計との差（▲ 上回る／▼ 下回る）"));
        Tooltip.install(
                kpiRemainingSub,
                new Tooltip("要 X m/日 = 予定合計に到達するために当日以降 1 日あたり必要な加工量\n= （予定合計 − 前日まで実績）÷ 残日数"));
        Tooltip.install(
                kpiProgressValue,
                new Tooltip(
                        "前日までの実績 ÷ 前日までの予定（参考値）。\n"
                                + "アラジン加工計画は完了した依頼が抽出から消えるため、前日までの予定は実際より少なく出ることがあります。\n"
                                + "前日までの予定が期間予定合計の 10% 未満のときは「—」にします。配台結果では算出しません。\n"
                                + String.format(Locale.ROOT, "%.0f%% 以上=緑、%.0f%% 以上=橙、未満=赤", PROGRESS_GOOD_PCT, PROGRESS_WARN_PCT)));
        renderEmpty("データ未読込", "タブを表示すると自動で読み込みます。読み込まれないときは「再読込」を押してください。");
    }

    public void bindShell(MainShellController shell) {
        this.shell = shell;
    }

    /** メインシェルでこのタブが選択されたとき。初回は読込、以降は指紋が変わったときだけ再読込。 */
    public void onMainShellTabSelected() {
        tabActive = true;
        updateAutoRefreshTimer();
        reloadFromSources(false);
    }

    /** 別タブへ移ったとき。自動更新を止める（再選択時に再開）。 */
    public void onMainShellTabDeselected() {
        tabActive = false;
        updateAutoRefreshTimer();
    }

    // ---- 初期化 ----------------------------------------------------------------------------

    private void initPeriodControls() {
        periodPresetCombo.setItems(FXCollections.observableArrayList(PeriodPreset.values()));
        // 「任意」は DatePicker を手で触った結果を示す表示専用。選択肢としては選べなくする
        periodPresetCombo.setCellFactory(
                lv ->
                        new ListCell<>() {
                            @Override
                            protected void updateItem(PeriodPreset item, boolean empty) {
                                super.updateItem(item, empty);
                                setText(empty || item == null ? null : item.toString());
                                setDisable(item == PeriodPreset.CUSTOM);
                            }
                        });
        periodPresetCombo.setButtonCell(
                new ListCell<>() {
                    @Override
                    protected void updateItem(PeriodPreset item, boolean empty) {
                        super.updateItem(item, empty);
                        setText(empty || item == null ? null : item.toString());
                    }
                });
        periodPresetCombo
                .getSelectionModel()
                .selectedItemProperty()
                .addListener(
                        (obs, o, n) -> {
                            if (suppressFilterEvents || n == null || n == PeriodPreset.CUSTOM) {
                                return;
                            }
                            clearInfoNotice();
                            applyPreset(n);
                            scheduleRecompute();
                        });
        fromDatePicker.valueProperty().addListener((obs, o, n) -> onDatePickerChanged(true));
        toDatePicker.valueProperty().addListener((obs, o, n) -> onDatePickerChanged(false));
        Tooltip.install(
                fromDatePicker,
                new Tooltip(
                        "期間開始（含む）。手で変えるとプリセットは「任意」になります。終了日より後にすると終了日も追従します"));
        Tooltip.install(
                toDatePicker,
                new Tooltip("期間終了（含む）。最長 " + MAX_PERIOD_DAYS + " 日。開始日より前にすると開始日も追従します"));
    }

    private void initAutoRefreshControls() {
        autoRefreshIntervalSpinner.setValueFactory(
                new SpinnerValueFactory.IntegerSpinnerValueFactory(
                        AUTO_REFRESH_MIN_SEC, AUTO_REFRESH_MAX_SEC, AUTO_REFRESH_DEFAULT_SEC, 30));
        autoRefreshIntervalSpinner.valueProperty().addListener((o, a, b) -> updateAutoRefreshTimer());
        autoRefreshIntervalSpinner
                .focusedProperty()
                .addListener(
                        (o, was, focused) -> {
                            if (!focused) {
                                commitSpinnerEditor();
                            }
                        });
        autoRefreshCheckBox.selectedProperty().addListener((o, a, b) -> updateAutoRefreshTimer());
        Tooltip.install(nextRefreshLabel, new Tooltip("自動更新の次回実行までの残り時間"));
        updateAutoRefreshTimer();
    }

    private void commitSpinnerEditor() {
        String text = autoRefreshIntervalSpinner.getEditor().getText();
        try {
            int v = Integer.parseInt(text.trim());
            autoRefreshIntervalSpinner
                    .getValueFactory()
                    .setValue(Math.max(AUTO_REFRESH_MIN_SEC, Math.min(AUTO_REFRESH_MAX_SEC, v)));
        } catch (NumberFormatException ex) {
            autoRefreshIntervalSpinner
                    .getEditor()
                    .setText(String.valueOf(autoRefreshIntervalSpinner.getValue()));
        }
    }

    private void initFilterCombos() {
        planSourceCombo.setItems(FXCollections.observableArrayList(PlanSource.values()));
        planSourceCombo.getSelectionModel().select(PlanSource.ALADDIN);
        planSourceCombo
                .getSelectionModel()
                .selectedItemProperty()
                .addListener((obs, o, n) -> onFilterChanged());
        machineCombo.setItems(FXCollections.observableArrayList(ALL_ITEM));
        machineCombo.getSelectionModel().selectFirst();
        machineCombo.getSelectionModel().selectedItemProperty().addListener((obs, o, n) -> onFilterChanged());
        processCombo.setItems(FXCollections.observableArrayList(ALL_ITEM));
        processCombo.getSelectionModel().selectFirst();
        processCombo.getSelectionModel().selectedItemProperty().addListener((obs, o, n) -> onFilterChanged());
        Tooltip.install(machineCombo, new Tooltip("機械名で絞込（実績・予定の両方に適用）"));
        Tooltip.install(processCombo, new Tooltip("工程名で絞込（実績・予定の両方に適用）"));
    }

    private void initViewModeToggles() {
        viewModeGroup
                .selectedToggleProperty()
                .addListener(
                        (obs, o, n) -> {
                            if (n == null && o != null) {
                                // 常に 1 つは選択状態にする
                                o.setSelected(true);
                                return;
                            }
                            applyViewMode(currentViewMode());
                        });
    }

    private void initCharts() {
        actualDailySeries.setName("実績");
        planDailySeries.setName("予定");
        actualCumSeries.setName("実績累計");
        planCumSeries.setName("予定累計");
        projectedCumSeries.setName("見込累計");
        dailyChart.getData().add(actualDailySeries);
        dailyChart.getData().add(planDailySeries);
        cumulativeChart.getData().add(actualCumSeries);
        cumulativeChart.getData().add(planCumSeries);
        cumulativeChart.getData().add(projectedCumSeries);
        cumulativeChart.setHorizontalZeroLineVisible(false);
        cumulativeChart.setVerticalZeroLineVisible(false);
        cumulativeChart.setPickOnBounds(false);
        // 目盛ラベルは自前描画（FXML: tickLabelsVisible=false）。ラベル行の高さは FXML の tickLabelGap で確保する。
        // prefHeight を固定してはいけない: CategoryAxis は setCategories 後 autoRanging=false になり、
        // カテゴリ間隔の再計算が computePrefHeight() 経由でしか走らないため、固定するとバーが左端に潰れる。
        StringConverter<Number> tickFmt =
                new StringConverter<>() {
                    @Override
                    public String toString(Number n) {
                        return n == null ? "" : formatM(n.doubleValue());
                    }

                    @Override
                    public Number fromString(String s) {
                        return null;
                    }
                };
        dailyYAxis.setTickLabelFormatter(tickFmt);
        cumulativeYAxis.setTickLabelFormatter(tickFmt);

        // 2 つのチャートのプロット領域を一致させる: 相手側 Y 軸の幅を自分の padding に載せる
        dailyYAxis.widthProperty().addListener((obs, o, n) -> syncChartPadding());
        cumulativeYAxis.widthProperty().addListener((obs, o, n) -> syncChartPadding());
        dailyChart.layoutBoundsProperty().addListener((obs, o, n) -> requestMarkerLayout());
        cumulativeChart.layoutBoundsProperty().addListener((obs, o, n) -> requestMarkerLayout());
        dailyXAxis.layoutBoundsProperty().addListener((obs, o, n) -> requestMarkerLayout());
        cumulativeXAxis.layoutBoundsProperty().addListener((obs, o, n) -> requestMarkerLayout());
        dailyYAxis.widthProperty().addListener((obs, o, n) -> requestMarkerLayout());
        cumulativeYAxis.widthProperty().addListener((obs, o, n) -> requestMarkerLayout());
        // CategoryAxis の座標は setCategories 後のレイアウトパスで categorySpacing が確定してから有効になる。
        // 軸のサイズが変わらずカテゴリ数だけ変わる（30 日→31 日など）ケースを拾う
        dailyXAxis.categorySpacingProperty().addListener((obs, o, n) -> requestMarkerLayout());
        cumulativeXAxis.categorySpacingProperty().addListener((obs, o, n) -> requestMarkerLayout());
        // タブ非選択中は detach されるので、再 attach（同サイズで bounds 変化なし）でも描き直す
        markerPane.sceneProperty().addListener((obs, o, n) -> {
            if (n != null) {
                requestMarkerLayout();
            }
        });

        todayLine.getStyleClass().add("pm-processing-trend-today-line");
        todayLine.setManaged(false);
        todayMarkerLabel.getStyleClass().add("pm-processing-trend-today-label");
        todayMarkerLabel.setManaged(false);
        todayLine.setVisible(false);
        todayMarkerLabel.setVisible(false);
        markerPane.getChildren().addAll(todayLine, todayMarkerLabel);
    }

    private void initDetailTable() {
        colDate.setCellValueFactory(cd -> new ReadOnlyObjectWrapper<>(cd.getValue()));
        colDate.setCellFactory(
                col ->
                        new TableCell<>() {
                            @Override
                            protected void updateItem(DayPoint item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty || item == null) {
                                    setText(null);
                                    pseudoClassStateChanged(PC_SAT, false);
                                    pseudoClassStateChanged(PC_SUN, false);
                                    return;
                                }
                                setText(formatDayWithWeekday(item.date()));
                                DayOfWeek dow = item.date().getDayOfWeek();
                                pseudoClassStateChanged(PC_SAT, dow == DayOfWeek.SATURDAY);
                                pseudoClassStateChanged(PC_SUN, dow == DayOfWeek.SUNDAY);
                            }
                        });
        colPlan.setCellValueFactory(cd -> new ReadOnlyObjectWrapper<>(cd.getValue().planM()));
        colActual.setCellValueFactory(cd -> new ReadOnlyObjectWrapper<>(cd.getValue().actualM()));
        // 当日以降は実績が揃っていないので差異を値として持たせない（セル側で「—」表示）
        colDiff.setCellValueFactory(
                cd ->
                        new ReadOnlyObjectWrapper<>(
                                cd.getValue().usesPlanForProjection() ? null : cd.getValue().diffM()));
        colPlanCum.setCellValueFactory(cd -> new ReadOnlyObjectWrapper<>(cd.getValue().planCumM()));
        colActualCum.setCellValueFactory(cd -> new ReadOnlyObjectWrapper<>(cd.getValue().actualCumM()));
        colProjectedCum.setCellValueFactory(
                cd -> new ReadOnlyObjectWrapper<>(cd.getValue().projectedCumM()));
        for (TableColumn<DayPoint, Number> c :
                List.of(colPlan, colActual, colPlanCum, colActualCum, colProjectedCum)) {
            c.setCellFactory(col -> numberCell(false));
            c.setStyle("-fx-alignment: CENTER-RIGHT;");
        }
        colDiff.setCellFactory(col -> numberCell(true));
        colDiff.setStyle("-fx-alignment: CENTER-RIGHT;");
        Tooltip.install(detailTable, new Tooltip("差異 = 実績 − 予定（当日以降は「—」）。見込累計 = 前日まで実績、当日以降は予定を採用"));
        detailTable.setRowFactory(
                tv ->
                        new TableRow<>() {
                            @Override
                            protected void updateItem(DayPoint item, boolean empty) {
                                super.updateItem(item, empty);
                                LocalDate today = currentResult != null ? currentResult.today() : LocalDate.now();
                                boolean has = !empty && item != null;
                                boolean weekend =
                                        has
                                                && (item.date().getDayOfWeek() == DayOfWeek.SATURDAY
                                                        || item.date().getDayOfWeek() == DayOfWeek.SUNDAY);
                                pseudoClassStateChanged(PC_WEEKEND, weekend);
                                pseudoClassStateChanged(PC_TODAY, has && item.date().equals(today));
                                pseudoClassStateChanged(PC_FUTURE, has && item.date().isAfter(today));
                            }
                        });
        detailTable.setPlaceholder(new Label("期間内のデータがありません"));
        detailTable.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_ALL_COLUMNS);
    }

    private TableCell<DayPoint, Number> numberCell(boolean signed) {
        return new TableCell<>() {
            @Override
            protected void updateItem(Number item, boolean empty) {
                super.updateItem(item, empty);
                if (empty) {
                    setText(null);
                    pseudoClassStateChanged(PC_GOOD, false);
                    pseudoClassStateChanged(PC_BAD, false);
                    return;
                }
                if (item == null) {
                    setText("—");
                    pseudoClassStateChanged(PC_GOOD, false);
                    pseudoClassStateChanged(PC_BAD, false);
                    return;
                }
                double v = item.doubleValue();
                setText(signed ? formatSigned(v) : formatM(v));
                pseudoClassStateChanged(PC_GOOD, signed && v > 0.5);
                pseudoClassStateChanged(PC_BAD, signed && v < -0.5);
            }
        };
    }

    private void initLegend() {
        legendBox.getChildren().clear();
        legendBox
                .getChildren()
                .addAll(
                        legendItem("pm-legend-swatch-bar-actual", "実績（日次）", ViewMode.CUMULATIVE),
                        legendItem("pm-legend-swatch-bar-plan", "予定（日次）", ViewMode.CUMULATIVE),
                        legendItem("pm-legend-swatch-line-actual", "実績累計", ViewMode.DAILY),
                        legendItem("pm-legend-swatch-line-plan", "予定累計", ViewMode.DAILY),
                        legendItem("pm-legend-swatch-line-projected", "見込累計", ViewMode.DAILY),
                        legendItem("pm-legend-swatch-today", "今日（前日まで実績／当日以降予定）", null));
    }

    /** @param hiddenIn このモードでは非表示にする（{@code null} なら常に表示） */
    private HBox legendItem(String swatchClass, String text, ViewMode hiddenIn) {
        Region swatch = new Region();
        swatch.getStyleClass().addAll("pm-legend-swatch", swatchClass);
        Label label = new Label(text);
        label.getStyleClass().add("pm-legend-label");
        HBox box = new HBox(5, swatch, label);
        box.setAlignment(Pos.CENTER_LEFT);
        if (hiddenIn != null) {
            box.getProperties().put("hiddenIn", hiddenIn);
        }
        return box;
    }

    private void initKeyboard() {
        tabRoot.setFocusTraversable(true);
        tabRoot.addEventFilter(KeyEvent.KEY_PRESSED, this::onKeyPressed);
    }

    private void onKeyPressed(KeyEvent e) {
        if (e.getCode() == KeyCode.F5 && !e.isControlDown() && !e.isAltDown() && !e.isMetaDown()) {
            reloadFromSources(true);
            e.consume();
            return;
        }
        if (!e.isShiftDown() || e.isControlDown() || e.isAltDown() || e.isMetaDown()) {
            return;
        }
        int dir = EquipmentStatusDashboardTabController.arrowDayShift(e.getCode());
        if (dir == 0 || EquipmentStatusDashboardTabController.skipDashboardShortcutTarget(e.getTarget())) {
            return;
        }
        shiftPeriod(dir);
        e.consume();
    }

    // ---- 期間 ----------------------------------------------------------------------------

    private void applyPreset(PeriodPreset preset) {
        LocalDate[] range = presetRange(preset, LocalDate.now());
        if (range == null) {
            return;
        }
        setPeriodSilently(range[0], range[1], preset);
    }

    /** プリセットの期間。CUSTOM は {@code null}。 */
    static LocalDate[] presetRange(PeriodPreset preset, LocalDate today) {
        return switch (preset) {
            case LAST_MONTH -> {
                LocalDate from = today.minusMonths(1).withDayOfMonth(1);
                yield new LocalDate[] {from, from.with(TemporalAdjusters.lastDayOfMonth())};
            }
            case NEXT_MONTH -> {
                LocalDate from = today.plusMonths(1).withDayOfMonth(1);
                yield new LocalDate[] {from, from.with(TemporalAdjusters.lastDayOfMonth())};
            }
            case AROUND_2_WEEKS -> new LocalDate[] {today.minusDays(13), today.plusDays(14)};
            case PAST_4_WEEKS -> new LocalDate[] {today.minusDays(27), today};
            case CUSTOM -> null;
            default -> {
                LocalDate from = today.withDayOfMonth(1);
                yield new LocalDate[] {from, from.with(TemporalAdjusters.lastDayOfMonth())};
            }
        };
    }

    private void setPeriodSilently(LocalDate from, LocalDate to, PeriodPreset preset) {
        boolean prev = suppressFilterEvents;
        suppressFilterEvents = true;
        try {
            fromDatePicker.setValue(from);
            toDatePicker.setValue(to);
            periodPresetCombo.getSelectionModel().select(preset);
        } finally {
            suppressFilterEvents = prev;
        }
    }

    /**
     * DatePicker を手で変えたとき。逆順（開始 &gt; 終了）なら変えた側に他方を追従させ、
     * 上限日数を超えたら終了日を詰める。
     */
    private void onDatePickerChanged(boolean fromChanged) {
        if (suppressFilterEvents) {
            return;
        }
        LocalDate from = fromDatePicker.getValue();
        LocalDate to = toDatePicker.getValue();
        if (from == null || to == null) {
            return;
        }
        clearInfoNotice();
        if (from.isAfter(to)) {
            if (fromChanged) {
                to = from;
            } else {
                from = to;
            }
        }
        long days = ChronoUnit.DAYS.between(from, to) + 1;
        if (days > MAX_PERIOD_DAYS) {
            to = from.plusDays(MAX_PERIOD_DAYS - 1L);
            showNotice(
                    "期間は最長 " + MAX_PERIOD_DAYS + " 日です。終了日を " + to.format(DATE_FMT) + " に調整しました。",
                    NoticeKind.INFO);
        }
        setPeriodSilently(from, to, presetMatching(from, to));
        scheduleRecompute();
    }

    @FXML
    private void onPrevPeriodAction() {
        shiftPeriod(-1);
    }

    @FXML
    private void onNextPeriodAction() {
        shiftPeriod(+1);
    }

    @FXML
    private void onThisMonthAction() {
        clearInfoNotice();
        applyPreset(PeriodPreset.THIS_MONTH);
        scheduleRecompute();
    }

    /** 月初〜月末なら月単位、それ以外は期間長ぶん日単位でずらす。 */
    private void shiftPeriod(int direction) {
        LocalDate from = fromDatePicker.getValue();
        LocalDate to = toDatePicker.getValue();
        clearInfoNotice();
        if (from == null || to == null) {
            applyPreset(PeriodPreset.THIS_MONTH);
            scheduleRecompute();
            return;
        }
        if (to.isBefore(from)) {
            LocalDate t = from;
            from = to;
            to = t;
        }
        LocalDate nf;
        LocalDate nt;
        boolean wholeMonths =
                from.getDayOfMonth() == 1 && to.equals(to.with(TemporalAdjusters.lastDayOfMonth()));
        if (wholeMonths) {
            long months = Math.max(1, ChronoUnit.MONTHS.between(from, to.plusDays(1)));
            nf = from.plusMonths(direction * months);
            nt = nf.plusMonths(months).minusDays(1);
        } else {
            long days = ChronoUnit.DAYS.between(from, to) + 1;
            nf = from.plusDays(direction * days);
            nt = to.plusDays(direction * days);
        }
        setPeriodSilently(nf, nt, presetMatching(nf, nt));
        scheduleRecompute();
    }

    private PeriodPreset presetMatching(LocalDate from, LocalDate to) {
        LocalDate today = LocalDate.now();
        for (PeriodPreset p : PeriodPreset.values()) {
            LocalDate[] r = presetRange(p, today);
            if (r != null && r[0].equals(from) && r[1].equals(to)) {
                return p;
            }
        }
        return PeriodPreset.CUSTOM;
    }

    /**
     * 日跨ぎ・月跨ぎ対応: プリセット選択中で、今日基準の期間が DatePicker とずれていたら追従する。
     *
     * @return 期間を動かしたら true
     */
    private boolean rolloverPeriodIfNeeded() {
        PeriodPreset preset = periodPresetCombo.getSelectionModel().getSelectedItem();
        if (preset == null || preset == PeriodPreset.CUSTOM) {
            return false;
        }
        LocalDate[] r = presetRange(preset, LocalDate.now());
        if (r == null) {
            return false;
        }
        if (r[0].equals(fromDatePicker.getValue()) && r[1].equals(toDatePicker.getValue())) {
            return false;
        }
        setPeriodSilently(r[0], r[1], preset);
        return true;
    }

    // ---- フィルタ・表示モード ---------------------------------------------------------------

    private void onFilterChanged() {
        if (suppressFilterEvents) {
            return;
        }
        scheduleRecompute();
    }

    private ViewMode currentViewMode() {
        if (viewModeGroup.getSelectedToggle() == viewDailyToggle) {
            return ViewMode.DAILY;
        }
        if (viewModeGroup.getSelectedToggle() == viewCumulativeToggle) {
            return ViewMode.CUMULATIVE;
        }
        return ViewMode.COMBO;
    }

    private void applyViewMode(ViewMode mode) {
        boolean showDaily = mode != ViewMode.CUMULATIVE;
        boolean showCum = mode != ViewMode.DAILY;
        dailyChart.setVisible(showDaily);
        dailyChart.setManaged(showDaily);
        cumulativeChart.setVisible(showCum);
        cumulativeChart.setManaged(showCum);
        boolean overlay = mode == ViewMode.COMBO;
        setStyleClassPresent(cumulativeChart, "pm-trend-overlay", overlay);
        cumulativeChart.setMouseTransparent(overlay);
        cumulativeChart.setHorizontalGridLinesVisible(!overlay);
        cumulativeYAxis.setSide(overlay ? Side.RIGHT : Side.LEFT);
        for (Node n : legendBox.getChildren()) {
            Object hiddenIn = n.getProperties().get("hiddenIn");
            boolean visible = hiddenIn == null || hiddenIn != mode;
            n.setVisible(visible);
            n.setManaged(visible);
        }
        if (chartTitleLabel != null) {
            chartTitleLabel.setText(
                    switch (mode) {
                        case DAILY -> "日次 実績・予定";
                        case CUMULATIVE -> "累計 実績・予定・見込";
                        default -> "加工実績・予定トレンド";
                    });
        }
        syncChartPadding();
        requestMarkerLayout();
    }

    private void syncChartPadding() {
        ViewMode mode = currentViewMode();
        double left = dailyYAxis.getWidth();
        double right = cumulativeYAxis.getWidth();
        if (mode == ViewMode.COMBO) {
            dailyChart.setPadding(new Insets(CHART_TOP_PADDING, right, 0, 0));
            cumulativeChart.setPadding(new Insets(CHART_TOP_PADDING, 0, 0, left));
        } else {
            dailyChart.setPadding(new Insets(CHART_TOP_PADDING, 0, 0, 0));
            cumulativeChart.setPadding(new Insets(CHART_TOP_PADDING, 0, 0, 0));
        }
    }

    // ---- 自動更新 --------------------------------------------------------------------------

    private int snapshotAutoRefreshIntervalSec() {
        Integer v = autoRefreshIntervalSpinner.getValue();
        if (v == null) {
            return AUTO_REFRESH_DEFAULT_SEC;
        }
        return Math.max(AUTO_REFRESH_MIN_SEC, Math.min(AUTO_REFRESH_MAX_SEC, v));
    }

    private void updateAutoRefreshTimer() {
        if (autoRefreshTimeline != null) {
            autoRefreshTimeline.stop();
            autoRefreshTimeline = null;
        }
        boolean want = tabActive && autoRefreshCheckBox.isSelected();
        if (!want) {
            nextRefreshLabel.setText(autoRefreshCheckBox.isSelected() ? "" : "自動更新 停止中");
            return;
        }
        autoRefreshRemainingSec = snapshotAutoRefreshIntervalSec();
        updateNextRefreshLabel();
        autoRefreshTimeline = new Timeline(new KeyFrame(Duration.seconds(1), e -> onAutoRefreshTick()));
        autoRefreshTimeline.setCycleCount(Timeline.INDEFINITE);
        autoRefreshTimeline.play();
    }

    private void onAutoRefreshTick() {
        autoRefreshRemainingSec--;
        if (autoRefreshRemainingSec <= 0) {
            autoRefreshRemainingSec = snapshotAutoRefreshIntervalSec();
            reloadFromSources(false);
        }
        updateNextRefreshLabel();
    }

    private void updateNextRefreshLabel() {
        int sec = Math.max(0, autoRefreshRemainingSec);
        nextRefreshLabel.setText(String.format(Locale.ROOT, "次回更新まで %d:%02d", sec / 60, sec % 60));
    }

    // ---- 読込 -------------------------------------------------------------------------------

    @FXML
    private void onReloadAction() {
        reloadFromSources(true);
    }

    /** 読込結果。機械名・工程名の抽出（全行走査・NFKC）は FX スレッドを塞がないよう Task 側で済ませる。 */
    private record ReloadOutcome(ReloadDecision decision, List<String> machines, List<String> processes) {}

    private void reloadFromSources(boolean userInitiated) {
        if (shell == null || reloadInFlight) {
            return;
        }
        reloadInFlight = true;
        setReloading(true);
        final SourceFingerprint previous = loadedFingerprint;
        final boolean haveCache = cachedSources != null;
        final Map<String, String> ui = shell.snapshotUiEnv();
        Task<ReloadOutcome> task =
                new Task<>() {
                    @Override
                    protected ReloadOutcome call() {
                        ReloadDecision decision =
                                EquipmentStatusDashboardSourceLoader.loadIfChanged(ui, previous, haveCache);
                        if (decision == null || decision.sourcesUnchanged() || decision.sources() == null) {
                            return new ReloadOutcome(decision, null, null);
                        }
                        LoadedSources s = decision.sources();
                        return new ReloadOutcome(
                                decision,
                                ProcessingTrendAggregator.machineNames(s.actuals(), s.aladdin(), s.dispatch()),
                                ProcessingTrendAggregator.processNames(s.actuals(), s.aladdin(), s.dispatch()));
                    }
                };
        task.setOnSucceeded(
                e -> {
                    reloadInFlight = false;
                    ReloadOutcome outcome = task.getValue();
                    ReloadDecision decision = outcome.decision();
                    lastSuccessAt = LocalTime.now();
                    lastCheckedAt = lastSuccessAt;
                    // 読込は成功したので、以前の読込エラーは指紋が変わっていなくても解消済み
                    lastLoadErrorDetail = "";
                    if (noticeKind == NoticeKind.LOAD_ERROR) {
                        hideNotice();
                    }
                    boolean periodMoved = rolloverPeriodIfNeeded();
                    if (decision == null || decision.sourcesUnchanged()) {
                        updateLastUpdatedLabel();
                        boolean dayRolled =
                                currentResult != null && !currentResult.today().equals(LocalDate.now());
                        setReloading(false);
                        if (currentResult == null || dayRolled || periodMoved) {
                            recomputeNow();
                        }
                        return;
                    }
                    lastDataChangedAt = lastSuccessAt;
                    loadedFingerprint = decision.fingerprint();
                    cachedSources = decision.sources();
                    updateLastUpdatedLabel();
                    updateSourceSummary();
                    try {
                        refreshFilterChoices(outcome.machines(), outcome.processes());
                    } finally {
                        setReloading(false);
                    }
                    recomputeNow();
                });
        task.setOnFailed(
                e -> {
                    reloadInFlight = false;
                    setReloading(false);
                    Throwable ex = task.getException();
                    String sourceContext = EquipmentStatusDashboardSourceLoader.formatSourceContext(ui);
                    lastLoadErrorDetail = DashboardLoadErrorFormatter.formatDetail(ex);
                    showNotice(
                            "読込に失敗しました"
                                    + (cachedSources != null && lastDataChangedAt != null
                                            ? "（" + lastDataChangedAt.format(TIME_FMT) + " 時点のデータを表示中）"
                                            : "")
                                    + " — "
                                    + lastLoadErrorDetail.replace('\n', ' '),
                            NoticeKind.LOAD_ERROR);
                    if (shell != null) {
                        shell.appendLog("[trend] 読込エラー: " + sourceContext);
                        shell.appendLog("[trend] " + lastLoadErrorDetail.replace('\n', ' '));
                    }
                    if (cachedSources == null) {
                        renderEmpty(
                                "ソースの読込に失敗しました",
                                "環境変数タブのソース設定を確認し「再読込」してください。");
                    }
                    if (userInitiated && shell != null) {
                        shell.showErrorDialog(
                                "加工トレンド 読込エラー",
                                sourceContext + "\n\n" + lastLoadErrorDetail);
                    }
                });
        pool.execute(task);
    }

    private void setReloading(boolean on) {
        reloadingFlag = on;
        loading.set(reloadingFlag || computeInFlight);
    }

    private void setComputing(boolean on) {
        computeInFlight = on;
        loading.set(reloadingFlag || computeInFlight);
    }

    private void refreshFilterChoices(List<String> machines, List<String> processes) {
        boolean prev = suppressFilterEvents;
        suppressFilterEvents = true;
        try {
            replaceComboItems(machineCombo, machines != null ? machines : List.of());
            replaceComboItems(processCombo, processes != null ? processes : List.of());
        } finally {
            suppressFilterEvents = prev;
        }
    }

    private static void replaceComboItems(ComboBox<String> combo, List<String> names) {
        String selected = combo.getSelectionModel().getSelectedItem();
        List<String> items = new ArrayList<>(names.size() + 1);
        items.add(ALL_ITEM);
        items.addAll(names);
        combo.setItems(FXCollections.observableArrayList(items));
        if (selected != null && items.contains(selected)) {
            combo.getSelectionModel().select(selected);
        } else {
            combo.getSelectionModel().selectFirst();
        }
    }

    // ---- 集計 -------------------------------------------------------------------------------

    private Filter currentFilter() {
        LocalDate from = fromDatePicker.getValue();
        LocalDate to = toDatePicker.getValue();
        if (from == null || to == null) {
            LocalDate today = LocalDate.now();
            from = today.withDayOfMonth(1);
            to = from.with(TemporalAdjusters.lastDayOfMonth());
        }
        PlanSource src = planSourceCombo.getSelectionModel().getSelectedItem();
        String machine = machineCombo.getSelectionModel().getSelectedItem();
        String process = processCombo.getSelectionModel().getSelectedItem();
        return new Filter(
                from,
                to,
                src,
                ALL_ITEM.equals(machine) ? null : machine,
                ALL_ITEM.equals(process) ? null : process);
    }

    /** UI 操作由来の再集計（デバウンスあり）。 */
    private void scheduleRecompute() {
        if (cachedSources == null) {
            return;
        }
        setComputing(true);
        recomputeDebounce.playFromStart();
    }

    private void recomputeNow() {
        recomputeDebounce.stop();
        if (cachedSources == null) {
            setComputing(false);
            return;
        }
        // 未開始のまま溜まっている旧集計は実行しない（世代チェックだけだと単一スレッド上でフル実行される）
        if (activeComputeTask != null) {
            activeComputeTask.cancel();
        }
        final LoadedSources src = cachedSources;
        final Filter filter = currentFilter();
        final LocalDate today = LocalDate.now();
        final long gen = computeGeneration.incrementAndGet();
        setComputing(true);
        Task<Result> task =
                new Task<>() {
                    @Override
                    protected Result call() {
                        return ProcessingTrendAggregator.aggregate(
                                src.actuals(), src.aladdin(), src.dispatch(), filter, today);
                    }
                };
        task.setOnSucceeded(
                e -> {
                    if (gen != computeGeneration.get()) {
                        return;
                    }
                    activeComputeTask = null;
                    setComputing(false);
                    currentResult = task.getValue();
                    render(currentResult, filter);
                });
        task.setOnFailed(
                e -> {
                    if (gen != computeGeneration.get()) {
                        return;
                    }
                    activeComputeTask = null;
                    setComputing(false);
                    Throwable ex = task.getException();
                    String detail = DashboardLoadErrorFormatter.formatDetail(ex);
                    showNotice("集計に失敗しました — " + detail.replace('\n', ' '), NoticeKind.COMPUTE_ERROR);
                    if (shell != null) {
                        shell.appendLog("[trend] 集計エラー: " + detail.replace('\n', ' '));
                    }
                });
        activeComputeTask = task;
        pool.execute(task);
    }

    // ---- 描画 -------------------------------------------------------------------------------

    private void render(Result r, Filter filter) {
        if (noticeKind == NoticeKind.COMPUTE_ERROR
                || (noticeKind == NoticeKind.LOAD_ERROR && lastLoadErrorDetail.isBlank())) {
            hideNotice();
        }
        renderKpis(r, filter);
        renderChart(r);
        detailTable.getItems().setAll(r.days());
        detailTable.refresh();
        renderDataWarnings(r);
        boolean empty = r.isEmpty();
        emptyStatePane.setVisible(empty);
        emptyStatePane.setManaged(empty);
        if (empty) {
            emptyStateTitle.setText("該当するデータがありません");
            emptyStateDetail.setText(
                    "期間 "
                            + filter.from().format(DATE_FMT)
                            + " 〜 "
                            + filter.to().format(DATE_FMT)
                            + "、機械="
                            + (filter.machine() != null ? filter.machine() : "すべて")
                            + "、工程="
                            + (filter.process() != null ? filter.process() : "すべて")
                            + "、予定="
                            + filter.planSource().label()
                            + " に実績・予定の行がありません。期間や絞込を見直してください。");
        }
        dailyChart.setOpacity(empty ? 0.25 : 1.0);
        cumulativeChart.setOpacity(empty ? 0.25 : 1.0);
        updateSourceSummary();
    }

    private void renderEmpty(String title, String detail) {
        currentResult = null;
        actualDailySeries.getData().clear();
        planDailySeries.getData().clear();
        actualCumSeries.getData().clear();
        planCumSeries.getData().clear();
        projectedCumSeries.getData().clear();
        detailTable.getItems().clear();
        for (Label l :
                List.of(kpiActualValue, kpiPlanValue, kpiProgressValue, kpiRemainingValue, kpiProjectedValue)) {
            l.setText("—");
        }
        for (Label l : List.of(kpiActualSub, kpiPlanSub, kpiProgressSub, kpiRemainingSub, kpiProjectedSub)) {
            l.setText("");
        }
        setProgressTone(kpiProgressValue, Double.NaN);
        setDiffTone(kpiProjectedSub, 0);
        emptyStateTitle.setText(title);
        emptyStateDetail.setText(detail);
        emptyStatePane.setVisible(true);
        emptyStatePane.setManaged(true);
        dailyChart.setOpacity(0.25);
        cumulativeChart.setOpacity(0.25);
        hideXAxisLabels();
        todayLine.setVisible(false);
        todayMarkerLabel.setVisible(false);
    }

    private void renderKpis(Result r, Filter filter) {
        LocalDate today = r.today();
        DayPoint todayPoint = null;
        int remainingDays = 0;
        for (DayPoint d : r.days()) {
            if (d.date().equals(today)) {
                todayPoint = d;
            }
            if (!d.date().isBefore(today)) {
                remainingDays++;
            }
        }
        boolean dispatchSource = filter.planSource() == PlanSource.DISPATCH;

        kpiActualValue.setText(formatM(r.actualTotalM()) + " m");
        kpiActualSub.setText(
                todayPoint != null
                        ? "うち当日 " + formatM(todayPoint.actualM()) + " m（前日まで " + formatM(r.actualToDateM()) + " m）"
                        : "");

        kpiPlanValue.setText(formatM(r.planTotalM()) + " m");
        if (dispatchSource) {
            // 配台結果は残量の先行き配台のみで、前日までの予定は構造的に持たない
            kpiPlanSub.setText("配台結果は当日以降の残量配台のみ");
        } else {
            kpiPlanSub.setText(todayPoint != null ? "前日まで " + formatM(r.planToDateM()) + " m" : "");
        }

        double pct = r.progressPct();
        boolean showPct = !Double.isNaN(pct) && !dispatchSource && r.progressDenominatorSufficient();
        kpiProgressValue.setText(showPct ? String.format(Locale.ROOT, "%.1f%%", pct) : "—");
        setProgressTone(kpiProgressValue, showPct ? pct : Double.NaN);
        if (dispatchSource) {
            kpiProgressSub.setText("配台結果では算出できません");
        } else if (Double.isNaN(pct)) {
            kpiProgressSub.setText("前日までの予定がありません");
        } else if (!showPct) {
            kpiProgressSub.setText("算出不可: 前日まで予定 " + formatM(r.planToDateM()) + " m のみ");
        } else {
            kpiProgressSub.setText(formatM(r.actualToDateM()) + " ÷ " + formatM(r.planToDateM()) + " m（参考値）");
        }

        kpiRemainingValue.setText(formatM(r.remainingPlanM()) + " m");
        if (remainingDays > 0) {
            double needPerDay = Math.max(0.0, r.planTotalM() - r.actualToDateM()) / remainingDays;
            kpiRemainingSub.setText("残 " + remainingDays + " 日・要 " + formatM(needPerDay) + " m/日");
        } else {
            kpiRemainingSub.setText("期間は終了しています");
        }

        kpiProjectedValue.setText(formatM(r.projectedTotalM()) + " m");
        double diff = r.projectedDiffM();
        String arrow = diff > 0.5 ? "▲ " : diff < -0.5 ? "▼ " : "";
        String ratio =
                r.planTotalM() > 0.5
                        ? String.format(Locale.ROOT, "（対予定 %+.1f%%）", diff / r.planTotalM() * 100.0)
                        : "";
        kpiProjectedSub.setText(arrow + formatSigned(diff) + " m" + ratio);
        setDiffTone(kpiProjectedSub, diff);
    }

    private void renderChart(Result r) {
        List<DayPoint> days = r.days();
        int n = days.size();
        boolean singleYear = n > 0 && days.get(0).date().getYear() == days.get(n - 1).date().getYear();
        boolean singleMonth =
                singleYear && days.get(0).date().getMonth() == days.get(n - 1).date().getMonth();
        // カテゴリ文字列は CategoryAxis 内で一意でなければならない（重複は IllegalArgumentException）。
        // 年をまたぐ期間では "9/3" が 2 回現れるので年を付ける
        DateTimeFormatter labelFmt =
                DateTimeFormatter.ofPattern(singleMonth ? "d" : singleYear ? "M/d" : "yy/M/d");
        List<String> labels = new ArrayList<>(n);
        List<LocalDate> dates = new ArrayList<>(n);
        for (DayPoint d : days) {
            labels.add(d.date().format(labelFmt));
            dates.add(d.date());
        }
        currentCategoryLabels = labels;
        currentDates = dates;
        Set<LocalDate> labelledDates = ProcessingTrendChartSupport.labelledDates(dates);
        Set<Integer> labelledIdx = new HashSet<>();
        for (int i = 0; i < n; i++) {
            if (labelledDates.contains(dates.get(i))) {
                labelledIdx.add(i);
            }
        }
        labelledCategoryIdx = labelledIdx;
        String axisLabel =
                n == 0
                        ? ""
                        : singleMonth
                                ? days.get(0).date().format(MONTH_FMT)
                                : days.get(0).date().format(DATE_FMT)
                                        + " 〜 "
                                        + days.get(n - 1).date().format(DATE_FMT);
        dailyXAxis.setLabel(axisLabel);
        cumulativeXAxis.setLabel(axisLabel);
        dailyXAxis.setCategories(FXCollections.observableArrayList(labels));
        cumulativeXAxis.setCategories(FXCollections.observableArrayList(labels));
        dailyChart.setCategoryGap(ProcessingTrendChartSupport.categoryGapFor(n));
        dailyChart.setBarGap(ProcessingTrendChartSupport.barGapFor(n));

        List<XYChart.Data<String, Number>> actDaily = new ArrayList<>(n);
        List<XYChart.Data<String, Number>> planDaily = new ArrayList<>(n);
        List<XYChart.Data<String, Number>> actCum = new ArrayList<>(n);
        List<XYChart.Data<String, Number>> planCum = new ArrayList<>(n);
        List<XYChart.Data<String, Number>> projCum = new ArrayList<>(n);
        double dailyMax = 0;
        double cumMax = 0;
        LocalDate projectedStart = r.today().minusDays(1);
        for (int i = 0; i < n; i++) {
            DayPoint d = days.get(i);
            String cat = labels.get(i);
            actDaily.add(new XYChart.Data<>(cat, d.actualM()));
            planDaily.add(new XYChart.Data<>(cat, d.planM()));
            dailyMax = Math.max(dailyMax, Math.max(d.actualM(), d.planM()));
            // 実績累計は当日までで線を止める（未来に水平線を伸ばさない）
            if (!d.date().isAfter(r.today())) {
                actCum.add(new XYChart.Data<>(cat, d.actualCumM()));
            }
            planCum.add(new XYChart.Data<>(cat, d.planCumM()));
            // 見込は実績の最終点（前日）から分岐させる。前日より前は実績累計と同一なので描かない
            if (!d.date().isBefore(projectedStart)) {
                projCum.add(new XYChart.Data<>(cat, d.projectedCumM()));
            }
            cumMax = Math.max(cumMax, Math.max(d.actualCumM(), Math.max(d.planCumM(), d.projectedCumM())));
        }
        applyNiceRange(dailyYAxis, dailyMax);
        applyNiceRange(cumulativeYAxis, cumMax);
        actualDailySeries.getData().setAll(actDaily);
        planDailySeries.getData().setAll(planDaily);
        actualCumSeries.getData().setAll(actCum);
        planCumSeries.getData().setAll(planCum);
        projectedCumSeries.getData().setAll(projCum);

        int projOffset = n - projCum.size();
        for (int i = 0; i < n; i++) {
            String tip = tooltipText(days.get(i), r.today());
            installSharedTooltip(actDaily.get(i), tip);
            installSharedTooltip(planDaily.get(i), tip);
            if (i < actCum.size()) {
                installSharedTooltip(actCum.get(i), tip);
            }
            installSharedTooltip(planCum.get(i), tip);
            if (i >= projOffset) {
                installSharedTooltip(projCum.get(i - projOffset), tip);
            }
        }
        requestMarkerLayout();
    }

    private static void applyNiceRange(ProcessingTrendNumberAxis axis, double max) {
        NiceRange nr = ProcessingTrendChartSupport.niceRange(max);
        axis.setLowerBound(0);
        axis.setUpperBound(nr.upperBound());
        axis.setTickUnit(nr.tickUnit());
    }

    private String tooltipText(DayPoint d, LocalDate today) {
        StringBuilder sb = new StringBuilder();
        sb.append(formatDayWithWeekday(d.date()));
        if (d.date().equals(today)) {
            sb.append("（今日）");
        }
        sb.append('\n');
        sb.append("実績 ").append(formatM(d.actualM())).append(" m");
        sb.append("  予定 ").append(formatM(d.planM())).append(" m");
        if (!d.usesPlanForProjection()) {
            sb.append("  差異 ").append(formatSigned(d.diffM())).append(" m");
        }
        sb.append('\n');
        sb.append("実績累計 ").append(formatM(d.actualCumM())).append(" m");
        sb.append("  予定累計 ").append(formatM(d.planCumM())).append(" m\n");
        String basis;
        if (!d.usesPlanForProjection()) {
            basis = "実績を採用";
        } else if (d.date().equals(today)) {
            basis = d.actualM() > d.planM() ? "当日: 実績が予定を上回るため実績を採用" : "当日: 予定を採用";
        } else {
            basis = "予定を採用";
        }
        sb.append("見込累計 ").append(formatM(d.projectedCumM())).append(" m（").append(basis).append("）");
        return sb.toString();
    }

    private void installSharedTooltip(XYChart.Data<String, Number> data, String text) {
        Node node = data.getNode();
        if (node != null) {
            bindSharedTooltip(node, text);
        }
        data.nodeProperty()
                .addListener(
                        (obs, o, nn) -> {
                            if (nn != null) {
                                bindSharedTooltip(nn, text);
                            }
                        });
    }

    private void bindSharedTooltip(Node node, String text) {
        node.setOnMouseEntered(e -> sharedChartTooltip.setText(text));
        Tooltip.install(node, sharedChartTooltip);
    }

    // ---- オーバーレイ（X 軸ラベル・今日マーカー） --------------------------------------------

    private void requestMarkerLayout() {
        if (markerLayoutScheduled) {
            return;
        }
        markerLayoutScheduled = true;
        Platform.runLater(
                () -> {
                    markerLayoutScheduled = false;
                    layoutOverlay();
                });
    }

    private void layoutOverlay() {
        Result r = currentResult;
        boolean chartShown = dailyChart.isVisible() || cumulativeChart.isVisible();
        if (r == null || r.days().isEmpty() || !chartShown) {
            hideXAxisLabels();
            todayLine.setVisible(false);
            todayMarkerLabel.setVisible(false);
            return;
        }
        CategoryAxis xAxis = dailyChart.isVisible() ? dailyXAxis : cumulativeXAxis;
        XYChart<String, Number> chart = dailyChart.isVisible() ? dailyChart : cumulativeChart;
        Node plotBg = chart.lookup(".chart-plot-background");
        if (plotBg == null || xAxis.getScene() == null || markerPane.getScene() == null) {
            hideXAxisLabels();
            todayLine.setVisible(false);
            todayMarkerLabel.setVisible(false);
            return;
        }
        layoutXAxisLabels(xAxis);
        layoutTodayMarker(r, xAxis, plotBg);
    }

    private void hideXAxisLabels() {
        for (Label l : xAxisLabelPool) {
            l.setVisible(false);
        }
    }

    private void layoutXAxisLabels(CategoryAxis xAxis) {
        int used = 0;
        double axisMinX = markerPane.sceneToLocal(xAxis.localToScene(0, 0)).getX();
        double axisMaxX = axisMinX + xAxis.getWidth();
        for (int i = 0; i < currentCategoryLabels.size(); i++) {
            if (!labelledCategoryIdx.contains(i)) {
                continue;
            }
            double ax = xAxis.getDisplayPosition(currentCategoryLabels.get(i));
            if (Double.isNaN(ax)) {
                continue;
            }
            Point2D p = markerPane.sceneToLocal(xAxis.localToScene(ax, X_LABEL_TOP_OFFSET));
            Label l = used < xAxisLabelPool.size() ? xAxisLabelPool.get(used) : newXAxisLabel();
            used++;
            l.setText(currentCategoryLabels.get(i));
            LocalDate d = currentDates.get(i);
            l.pseudoClassStateChanged(PC_SAT, d.getDayOfWeek() == DayOfWeek.SATURDAY);
            l.pseudoClassStateChanged(PC_SUN, d.getDayOfWeek() == DayOfWeek.SUNDAY);
            l.applyCss();
            l.autosize();
            double w = l.getWidth();
            double lx = p.getX() - w / 2.0;
            // 端のラベルは軸の外へはみ出さない
            lx = Math.max(axisMinX, Math.min(lx, axisMaxX - w));
            l.relocate(Math.round(lx), Math.round(p.getY()));
            l.setVisible(true);
        }
        for (int i = used; i < xAxisLabelPool.size(); i++) {
            xAxisLabelPool.get(i).setVisible(false);
        }
    }

    private Label newXAxisLabel() {
        Label l = new Label();
        l.getStyleClass().add("pm-processing-trend-x-label");
        l.setManaged(false);
        l.setMouseTransparent(true);
        xAxisLabelPool.add(l);
        markerPane.getChildren().add(0, l);
        return l;
    }

    private void layoutTodayMarker(Result r, CategoryAxis xAxis, Node plotBg) {
        LocalDate today = r.today();
        LocalDate first = r.days().get(0).date();
        LocalDate last = r.days().get(r.days().size() - 1).date();
        if (today.isBefore(first) || today.isAfter(last.plusDays(1))) {
            todayLine.setVisible(false);
            todayMarkerLabel.setVisible(false);
            return;
        }
        double x;
        if (today.isAfter(last)) {
            // 期間の翌日 = 期間全体が実績側。右端に線を置く
            int idx = currentCategoryLabels.size() - 1;
            x = xAxis.getDisplayPosition(currentCategoryLabels.get(idx)) + xAxis.getCategorySpacing() / 2.0;
        } else {
            int idx = (int) ChronoUnit.DAYS.between(first, today);
            if (idx < 0 || idx >= currentCategoryLabels.size()) {
                todayLine.setVisible(false);
                todayMarkerLabel.setVisible(false);
                return;
            }
            x = xAxis.getDisplayPosition(currentCategoryLabels.get(idx)) - xAxis.getCategorySpacing() / 2.0;
        }
        Point2D top = markerPane.sceneToLocal(xAxis.localToScene(x, 0));
        var plotBounds = markerPane.sceneToLocal(plotBg.localToScene(plotBg.getBoundsInLocal()));
        if (Double.isNaN(top.getX()) || plotBounds == null) {
            todayLine.setVisible(false);
            todayMarkerLabel.setVisible(false);
            return;
        }
        double px = Math.round(top.getX()) + 0.5;
        todayLine.setStartX(px);
        todayLine.setEndX(px);
        todayLine.setStartY(plotBounds.getMinY());
        todayLine.setEndY(plotBounds.getMaxY());
        todayLine.setVisible(true);
        todayMarkerLabel.applyCss();
        todayMarkerLabel.autosize();
        double labelW = todayMarkerLabel.getWidth();
        // 右端に寄ったらラベルを線の左側へ
        double lx = px + 4 + labelW > plotBounds.getMaxX() ? px - 4 - labelW : px + 4;
        todayMarkerLabel.relocate(lx, plotBounds.getMinY() + 2);
        todayMarkerLabel.setVisible(true);
    }

    // ---- 表示補助 --------------------------------------------------------------------------

    private void applyLoadingPresentation(boolean isLoading) {
        loadingIndicator.setVisible(isLoading);
        loadingIndicator.setManaged(isLoading);
        loadingStatusLabel.setVisible(isLoading);
        loadingStatusLabel.setManaged(isLoading);
        reloadButton.setDisable(isLoading);
        if (noticeReloadButton != null) {
            noticeReloadButton.setDisable(isLoading);
        }
        if (isLoading && currentResult == null) {
            emptyStateTitle.setText("読込中…");
            emptyStateDetail.setText("加工実績・加工計画・配台結果を読み込んでいます。");
            emptyStatePane.setVisible(true);
            emptyStatePane.setManaged(true);
        }
    }

    private void updateLastUpdatedLabel() {
        StringBuilder sb = new StringBuilder("最終更新: ");
        sb.append(lastDataChangedAt != null ? lastDataChangedAt.format(TIME_FMT) : "—");
        if (lastCheckedAt != null && (lastDataChangedAt == null || !lastCheckedAt.equals(lastDataChangedAt))) {
            sb.append("（確認 ").append(lastCheckedAt.format(TIME_FMT)).append("）");
        }
        lastUpdatedLabel.setText(sb.toString());
    }

    private void updateSourceSummary() {
        if (cachedSources == null) {
            sourceSummaryLabel.setText("");
            sourceSummaryLabel.setVisible(false);
            sourceSummaryLabel.setManaged(false);
            return;
        }
        StringBuilder sb = new StringBuilder();
        sb.append("実績=").append(cachedSources.actualSourceLabel());
        sb.append("  アラジン=").append(cachedSources.aladdinSourceLabel());
        sb.append("  配台=").append(cachedSources.dispatchSourceLabel());
        String stats = EquipmentStatusDashboardSourceLoader.formatLoadStatsSummary(cachedSources.loadStats());
        if (stats != null && !stats.isBlank()) {
            sb.append("  ").append(stats);
        }
        if (currentResult != null) {
            sb.append("  集計対象: 実績 ")
                    .append(currentResult.actualRowsCounted())
                    .append(" 行・予定 ")
                    .append(currentResult.planRowsCounted())
                    .append(" 行");
            if (currentResult.actualMinDate() != null && currentResult.actualMaxDate() != null) {
                sb.append("  実績加工日 ")
                        .append(currentResult.actualMinDate().format(DATE_FMT))
                        .append("〜")
                        .append(currentResult.actualMaxDate().format(DATE_FMT));
            }
        }
        if (cachedSources.loadNotice() != null && !cachedSources.loadNotice().isBlank()) {
            sb.append("  ※ ").append(cachedSources.loadNotice());
        }
        sourceSummaryLabel.setText(sb.toString());
        boolean show = !sb.isEmpty();
        sourceSummaryLabel.setVisible(show);
        sourceSummaryLabel.setManaged(show);
    }

    private void showNotice(String text, NoticeKind kind) {
        noticeKind = kind;
        boolean error = kind == NoticeKind.LOAD_ERROR || kind == NoticeKind.COMPUTE_ERROR;
        setStyleClassPresent(noticeBanner, "pm-processing-trend-banner-error", error);
        setStyleClassPresent(noticeBanner, "pm-processing-trend-banner-warn", !error);
        if (noticeReloadButton != null) {
            noticeReloadButton.setVisible(error);
            noticeReloadButton.setManaged(error);
        }
        noticeBannerLabel.setText(text);
        noticeBanner.setVisible(true);
        noticeBanner.setManaged(true);
    }

    private void hideNotice() {
        noticeKind = NoticeKind.NONE;
        noticeBanner.setVisible(false);
        noticeBanner.setManaged(false);
    }

    /** 集計結果に付随する注意をバナーに出す。読込エラー表示中は上書きしない。 */
    private void renderDataWarnings(Result r) {
        List<String> messages = new ArrayList<>(r.warnings());
        if (r.periodStartsBeforeActualSource()) {
            messages.add(
                    "実績ソースに含まれる加工日は "
                            + r.actualMinDate().format(DATE_FMT)
                            + " 以降のみです。それより前の実績 0 はデータが無いことを意味します。");
        }
        if (messages.isEmpty()) {
            if (noticeKind == NoticeKind.DATA_WARN) {
                hideNotice();
            }
            return;
        }
        if (noticeKind == NoticeKind.LOAD_ERROR && !lastLoadErrorDetail.isBlank()) {
            return;
        }
        showNotice(String.join("　", messages), NoticeKind.DATA_WARN);
    }

    private void clearInfoNotice() {
        if (noticeKind == NoticeKind.INFO) {
            hideNotice();
        }
    }

    private static void setProgressTone(Label label, double pct) {
        boolean has = !Double.isNaN(pct);
        label.pseudoClassStateChanged(PC_GOOD, has && pct >= PROGRESS_GOOD_PCT);
        label.pseudoClassStateChanged(PC_WARN, has && pct >= PROGRESS_WARN_PCT && pct < PROGRESS_GOOD_PCT);
        label.pseudoClassStateChanged(PC_BAD, has && pct < PROGRESS_WARN_PCT);
    }

    private static void setDiffTone(Label label, double diff) {
        label.pseudoClassStateChanged(PC_GOOD, diff > 0.5);
        label.pseudoClassStateChanged(PC_BAD, diff < -0.5);
    }

    private static void setStyleClassPresent(Node node, String styleClass, boolean present) {
        boolean has = node.getStyleClass().contains(styleClass);
        if (present && !has) {
            node.getStyleClass().add(styleClass);
        } else if (!present && has) {
            node.getStyleClass().remove(styleClass);
        }
    }

    static String formatM(double v) {
        if (Math.abs(v) < 0.05) {
            return "0";
        }
        if (Math.abs(v) < 100.0 && Math.abs(v - Math.rint(v)) > 0.05) {
            return String.format(Locale.JAPAN, "%,.1f", v);
        }
        return String.format(Locale.JAPAN, "%,.0f", v);
    }

    static String formatSigned(double v) {
        String body = formatM(Math.abs(v));
        if (Math.abs(v) < 0.05) {
            return "±0";
        }
        return (v > 0 ? "+" : "−") + body;
    }

    static String formatDayWithWeekday(LocalDate d) {
        Objects.requireNonNull(d);
        return d.format(DateTimeFormatter.ofPattern("M/d"))
                + "("
                + d.getDayOfWeek().getDisplayName(TextStyle.SHORT, Locale.JAPAN)
                + ")";
    }
}
