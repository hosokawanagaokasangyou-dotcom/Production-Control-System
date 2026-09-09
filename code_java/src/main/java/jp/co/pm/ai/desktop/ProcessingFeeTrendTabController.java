package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.text.NumberFormat;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.time.format.TextStyle;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.atomic.AtomicLong;

import javafx.application.Platform;
import javafx.beans.property.ReadOnlyObjectWrapper;
import javafx.collections.FXCollections;
import javafx.concurrent.Task;
import javafx.css.PseudoClass;
import javafx.fxml.FXML;
import javafx.geometry.Insets;
import javafx.geometry.Point2D;
import javafx.geometry.Side;
import javafx.scene.Node;
import javafx.scene.chart.BarChart;
import javafx.scene.chart.CategoryAxis;
import javafx.scene.chart.LineChart;
import javafx.scene.chart.XYChart;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.DatePicker;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableRow;
import javafx.scene.control.TableView;
import javafx.scene.control.TitledPane;
import javafx.scene.control.Tooltip;
import javafx.scene.layout.Background;
import javafx.scene.layout.BackgroundFill;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.CornerRadii;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Pane;
import javafx.scene.layout.Region;
import javafx.scene.layout.StackPane;
import javafx.scene.paint.Color;
import javafx.scene.shape.Line;
import javafx.scene.shape.LineTo;
import javafx.scene.shape.MoveTo;
import javafx.scene.shape.Path;
import javafx.util.StringConverter;

import jp.co.pm.ai.desktop.ProcessingTrendChartSupport.NiceRange;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardSourceLoader;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardSourceLoader.LoadedSources;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardSourceLoader.ReloadDecision;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardSourceLoader.SourceFingerprint;
import jp.co.pm.ai.desktop.io.actuals.JuchuProcessingFeeRateLoader;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.DayPoint;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.QuantityLine;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.RequestPoint;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.Result;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendQuantityExtractor;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.ActualSource;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.Filter;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.PlanSource;

/**
 * 「加工賃」子タブ: AH×m の実績円・予定円・累計折れ線（加工量トレンドと同系の軸デザイン）。
 */
public class ProcessingFeeTrendTabController {

    private static final double CHART_TOP_PADDING = 6.0;
    private static final double X_LABEL_TOP_OFFSET = 7.0;
    private static final Background CHART_STACK_WHITE =
            new Background(new BackgroundFill(Color.WHITE, new CornerRadii(4), Insets.EMPTY));
    private static final Background PLOT_WHITE =
            new Background(new BackgroundFill(Color.WHITE, CornerRadii.EMPTY, Insets.EMPTY));
    private static final Background PLOT_TRANSPARENT =
            new Background(new BackgroundFill(Color.TRANSPARENT, CornerRadii.EMPTY, Insets.EMPTY));
    private static final PseudoClass PC_SAT = PseudoClass.getPseudoClass("sat");
    private static final PseudoClass PC_SUN = PseudoClass.getPseudoClass("sun");
    private static final PseudoClass PC_WEEKEND = PseudoClass.getPseudoClass("weekend");
    private static final PseudoClass PC_TODAY = PseudoClass.getPseudoClass("today");
    private static final PseudoClass PC_FUTURE = PseudoClass.getPseudoClass("future");
    private static final PseudoClass PC_GOOD = PseudoClass.getPseudoClass("good");
    private static final PseudoClass PC_BAD = PseudoClass.getPseudoClass("bad");

    private enum PeriodPreset {
        THIS_MONTH("今月"),
        LAST_MONTH("先月"),
        PAST_3_MONTHS("直近3ヶ月"),
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

    @FXML private BorderPane tabRoot;
    @FXML private Button reloadButton;
    @FXML private ComboBox<PeriodPreset> periodPresetCombo;
    @FXML private DatePicker fromDatePicker;
    @FXML private DatePicker toDatePicker;
    @FXML private Button thisMonthButton;
    @FXML private HBox loadingChip;
    @FXML private ProgressIndicator loadingIndicator;
    @FXML private Label loadingStatusLabel;
    @FXML private ComboBox<ActualSource> actualSourceCombo;
    @FXML private ComboBox<PlanSource> planSourceCombo;
    @FXML private Label kpiActualYen;
    @FXML private Label kpiPlanYen;
    @FXML private HBox noticeBanner;
    @FXML private Label noticeLabel;
    @FXML private Label chartTitleLabel;
    @FXML private HBox legendBox;
    @FXML private StackPane chartStack;
    @FXML private BarChart<String, Number> dailyChart;
    @FXML private CategoryAxis dailyXAxis;
    @FXML private ProcessingTrendNumberAxis dailyYAxis;
    @FXML private LineChart<String, Number> cumulativeChart;
    @FXML private CategoryAxis cumulativeXAxis;
    @FXML private ProcessingTrendNumberAxis cumulativeYAxis;
    @FXML private Pane markerPane;
    @FXML private Label sourceSummaryLabel;
    @FXML private TitledPane detailPane;
    @FXML private TableView<DayPoint> detailTable;
    @FXML private TableColumn<DayPoint, DayPoint> colDate;
    @FXML private TableColumn<DayPoint, Number> colActualYen;
    @FXML private TableColumn<DayPoint, Number> colPlanYen;
    @FXML private TableColumn<DayPoint, Number> colDiffYen;
    @FXML private TableColumn<DayPoint, Number> colActualCumYen;
    @FXML private TableColumn<DayPoint, Number> colPlanCumYen;
    @FXML private TitledPane requestPane;
    @FXML private TableView<RequestPoint> requestTable;
    @FXML private TableColumn<RequestPoint, String> colRequestNo;
    @FXML private TableColumn<RequestPoint, Number> colRateYen;
    @FXML private TableColumn<RequestPoint, Number> colActualM;
    @FXML private TableColumn<RequestPoint, Number> colPlanM;
    @FXML private TableColumn<RequestPoint, Number> colReqActualYen;
    @FXML private TableColumn<RequestPoint, Number> colReqPlanYen;
    @FXML private TableColumn<RequestPoint, Number> colReqDiffYen;

    private MainShellController shell;
    private boolean suppressFilterEvents;
    private boolean tabActive;
    private boolean reloadInFlight;
    private LoadedSources cachedSources;
    private SourceFingerprint loadedFingerprint;
    private Map<String, Double> feeRates = Map.of();
    private Result currentResult;
    private List<String> currentCategoryLabels = List.of();
    private List<LocalDate> currentDates = List.of();
    private Set<Integer> labelledCategoryIdx = Set.of();
    private final AtomicLong computeSeq = new AtomicLong();
    private final List<Label> xAxisLabelPool = new ArrayList<>();
    private boolean overlayLayoutScheduled;
    private final Line todayLine = new Line();
    private final Label todayMarkerLabel = new Label("今日");

    private final XYChart.Series<String, Number> actualSeries = new XYChart.Series<>();
    private final XYChart.Series<String, Number> planSeries = new XYChart.Series<>();
    /** LineChart は第二軸シェル専用（系列データは載せない。折れ線は Path）。 */
    private final XYChart.Series<String, Number> actualCumSeries = new XYChart.Series<>();
    private final XYChart.Series<String, Number> planCumSeries = new XYChart.Series<>();
    private final Path actualCumPath = new Path();
    private final Path planCumPath = new Path();
    private OverlayPolyline overlayActualCum = OverlayPolyline.EMPTY;
    private OverlayPolyline overlayPlanCum = OverlayPolyline.EMPTY;

    /** markerPane 上の累計折れ線用。 */
    private record OverlayPolyline(List<String> categories, List<Double> values) {
        static final OverlayPolyline EMPTY = new OverlayPolyline(List.of(), List.of());

        static OverlayPolyline fromChartData(List<XYChart.Data<String, Number>> data) {
            if (data == null || data.isEmpty()) {
                return EMPTY;
            }
            List<String> cats = new ArrayList<>(data.size());
            List<Double> vals = new ArrayList<>(data.size());
            for (XYChart.Data<String, Number> d : data) {
                if (d == null || d.getXValue() == null || d.getYValue() == null) {
                    continue;
                }
                cats.add(d.getXValue());
                vals.add(d.getYValue().doubleValue());
            }
            return cats.isEmpty() ? EMPTY : new OverlayPolyline(List.copyOf(cats), List.copyOf(vals));
        }

        boolean isEmpty() {
            return categories.isEmpty();
        }
    }

    @FXML
    private void initialize() {
        actualSeries.setName("実績円");
        planSeries.setName("予定円");
        dailyChart.getData().setAll(actualSeries, planSeries);
        dailyChart.setAnimated(false);

        actualCumSeries.setName("実績累計");
        planCumSeries.setName("予定累計");
        cumulativeChart.setAnimated(false);
        cumulativeChart.setCreateSymbols(false);
        cumulativeChart.setAxisSortingPolicy(LineChart.SortingPolicy.NONE);
        cumulativeChart.setHorizontalZeroLineVisible(false);
        cumulativeChart.setVerticalZeroLineVisible(false);
        cumulativeChart.setPickOnBounds(false);
        // 系列に累計データを載せ右軸を立てる。棒は dailyChart（左軸）。
        cumulativeChart.getData().setAll(actualCumSeries, planCumSeries);
        if (!cumulativeChart.getStyleClass().contains("pm-trend-overlay")) {
            cumulativeChart.getStyleClass().add("pm-trend-overlay");
        }
        cumulativeYAxis.setSide(Side.RIGHT);
        cumulativeYAxis.setTickLabelsVisible(true);
        cumulativeYAxis.setLabel("累計 (円) ─ 折れ線");
        cumulativeYAxis.setMinWidth(Region.USE_COMPUTED_SIZE);
        cumulativeChart.setMouseTransparent(true);
        cumulativeChart.setHorizontalGridLinesVisible(false);

        StringConverter<Number> tickFmt =
                new StringConverter<>() {
                    @Override
                    public String toString(Number n) {
                        return n == null ? "" : formatAxisTick(n.doubleValue());
                    }

                    @Override
                    public Number fromString(String s) {
                        return null;
                    }
                };
        dailyYAxis.setAutoRanging(false);
        cumulativeYAxis.setAutoRanging(false);
        dailyYAxis.setTickLabelFormatter(tickFmt);
        cumulativeYAxis.setTickLabelFormatter(tickFmt);

        initCumOverlayPaths();
        initTodayMarker();
        initLegend();
        applyChartWhiteBackground();
        // スキン生成後にも白地を再適用（テーマの plot 塗りつぶしを上書き）
        chartStack.sceneProperty().addListener((o, a, n) -> Platform.runLater(this::applyChartWhiteBackground));
        dailyChart.layoutBoundsProperty().addListener((o, a, n) -> applyChartWhiteBackground());
        cumulativeChart.layoutBoundsProperty().addListener((o, a, n) -> applyChartWhiteBackground());

        periodPresetCombo.setItems(
                FXCollections.observableArrayList(
                        PeriodPreset.THIS_MONTH,
                        PeriodPreset.LAST_MONTH,
                        PeriodPreset.PAST_3_MONTHS,
                        PeriodPreset.CUSTOM));
        actualSourceCombo.setItems(FXCollections.observableArrayList(ActualSource.values()));
        planSourceCombo.setItems(FXCollections.observableArrayList(PlanSource.values()));
        actualSourceCombo.getSelectionModel().select(ActualSource.DAILY_REPORT);
        planSourceCombo.getSelectionModel().select(PlanSource.ALADDIN);

        suppressFilterEvents = true;
        applyPreset(PeriodPreset.THIS_MONTH);
        periodPresetCombo.getSelectionModel().select(PeriodPreset.THIS_MONTH);
        suppressFilterEvents = false;

        periodPresetCombo
                .getSelectionModel()
                .selectedItemProperty()
                .addListener(
                        (o, a, n) -> {
                            if (suppressFilterEvents || n == null || n == PeriodPreset.CUSTOM) {
                                return;
                            }
                            applyPreset(n);
                            scheduleRecompute();
                        });
        fromDatePicker.valueProperty().addListener((o, a, n) -> onDateManual());
        toDatePicker.valueProperty().addListener((o, a, n) -> onDateManual());
        actualSourceCombo.valueProperty().addListener((o, a, n) -> scheduleRecompute());
        planSourceCombo.valueProperty().addListener((o, a, n) -> scheduleRecompute());

        dailyYAxis.widthProperty().addListener((o, a, n) -> syncChartPadding());
        cumulativeYAxis.widthProperty().addListener((o, a, n) -> syncChartPadding());
        dailyChart.layoutBoundsProperty().addListener((o, a, n) -> requestOverlayLayout());
        cumulativeChart.layoutBoundsProperty().addListener((o, a, n) -> requestOverlayLayout());
        dailyXAxis.layoutBoundsProperty().addListener((o, a, n) -> requestOverlayLayout());
        if (cumulativeXAxis != null) {
            cumulativeXAxis.layoutBoundsProperty().addListener((o, a, n) -> requestOverlayLayout());
        }
        dailyYAxis.widthProperty().addListener((o, a, n) -> requestOverlayLayout());
        cumulativeYAxis.widthProperty().addListener((o, a, n) -> requestOverlayLayout());
        dailyXAxis.categorySpacingProperty().addListener((o, a, n) -> requestOverlayLayout());
        markerPane.sceneProperty().addListener((o, a, n) -> requestOverlayLayout());

        sourceSummaryLabel.setText("受注 AH × 加工日 m。工程延べのため依頼生産量とは一致しません。");
        initDetailTable();
        initRequestTable();
        syncChartPadding();
        renderEmpty();
    }

    private void initDetailTable() {
        if (detailTable == null) {
            return;
        }
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
                                LocalDate d = item.date();
                                setText(formatDayWithWeekday(d));
                                DayOfWeek dow = d.getDayOfWeek();
                                pseudoClassStateChanged(PC_SAT, dow == DayOfWeek.SATURDAY);
                                pseudoClassStateChanged(PC_SUN, dow == DayOfWeek.SUNDAY);
                            }
                        });
        colActualYen.setCellValueFactory(cd -> new ReadOnlyObjectWrapper<>(cd.getValue().actualYen()));
        colPlanYen.setCellValueFactory(cd -> new ReadOnlyObjectWrapper<>(cd.getValue().planYen()));
        colDiffYen.setCellValueFactory(
                cd -> new ReadOnlyObjectWrapper<>(cd.getValue().actualYen() - cd.getValue().planYen()));
        colActualCumYen.setCellValueFactory(
                cd -> new ReadOnlyObjectWrapper<>(cd.getValue().actualCumYen()));
        colPlanCumYen.setCellValueFactory(
                cd -> new ReadOnlyObjectWrapper<>(cd.getValue().planCumYen()));
        for (TableColumn<DayPoint, Number> c :
                List.of(colActualYen, colPlanYen, colActualCumYen, colPlanCumYen)) {
            c.setCellFactory(col -> yenCell(false));
            c.setStyle("-fx-alignment: CENTER-RIGHT;");
        }
        colDiffYen.setCellFactory(col -> yenCell(true));
        colDiffYen.setStyle("-fx-alignment: CENTER-RIGHT;");
        detailTable.setRowFactory(
                tv ->
                        new TableRow<>() {
                            @Override
                            protected void updateItem(DayPoint item, boolean empty) {
                                super.updateItem(item, empty);
                                LocalDate today =
                                        currentResult != null ? currentResult.today() : LocalDate.now();
                                if (empty || item == null) {
                                    pseudoClassStateChanged(PC_WEEKEND, false);
                                    pseudoClassStateChanged(PC_TODAY, false);
                                    pseudoClassStateChanged(PC_FUTURE, false);
                                    return;
                                }
                                DayOfWeek dow = item.date().getDayOfWeek();
                                boolean weekend =
                                        dow == DayOfWeek.SATURDAY || dow == DayOfWeek.SUNDAY;
                                pseudoClassStateChanged(PC_WEEKEND, weekend);
                                pseudoClassStateChanged(PC_TODAY, item.date().equals(today));
                                pseudoClassStateChanged(PC_FUTURE, item.date().isAfter(today));
                            }
                        });
        detailTable.setPlaceholder(new Label("期間内のデータがありません"));
        detailTable.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_ALL_COLUMNS);
        Tooltip.install(
                detailTable,
                new Tooltip("差異 = 実績円 − 予定円。単位は円（AH × 工程延べ m）。"));
    }

    private void initRequestTable() {
        if (requestTable == null) {
            return;
        }
        colRequestNo.setCellValueFactory(
                cd -> new ReadOnlyObjectWrapper<>(cd.getValue().requestNo()));
        colRateYen.setCellValueFactory(
                cd ->
                        new ReadOnlyObjectWrapper<>(
                                cd.getValue().rateMissing()
                                        ? null
                                        : cd.getValue().rateYenPerM()));
        colRateYen.setCellFactory(
                col ->
                        new TableCell<>() {
                            @Override
                            protected void updateItem(Number item, boolean empty) {
                                super.updateItem(item, empty);
                                pseudoClassStateChanged(PC_BAD, false);
                                if (empty) {
                                    setText(null);
                                    return;
                                }
                                RequestPoint row = getTableRow() != null ? getTableRow().getItem() : null;
                                if (row != null && row.rateMissing()) {
                                    setText("—");
                                    pseudoClassStateChanged(PC_BAD, true);
                                    return;
                                }
                                if (item == null) {
                                    setText("—");
                                    return;
                                }
                                setText(
                                        NumberFormat.getIntegerInstance(Locale.JAPAN)
                                                .format(Math.rint(item.doubleValue())));
                            }
                        });
        colRateYen.setStyle("-fx-alignment: CENTER-RIGHT;");
        colActualM.setCellValueFactory(
                cd -> new ReadOnlyObjectWrapper<>(cd.getValue().actualMeters()));
        colPlanM.setCellValueFactory(
                cd -> new ReadOnlyObjectWrapper<>(cd.getValue().planMeters()));
        for (TableColumn<RequestPoint, Number> c : List.of(colActualM, colPlanM)) {
            c.setCellFactory(col -> metersCell());
            c.setStyle("-fx-alignment: CENTER-RIGHT;");
        }
        colReqActualYen.setCellValueFactory(
                cd -> new ReadOnlyObjectWrapper<>(cd.getValue().actualYen()));
        colReqPlanYen.setCellValueFactory(
                cd -> new ReadOnlyObjectWrapper<>(cd.getValue().planYen()));
        colReqDiffYen.setCellValueFactory(
                cd ->
                        new ReadOnlyObjectWrapper<>(
                                cd.getValue().actualYen() - cd.getValue().planYen()));
        for (TableColumn<RequestPoint, Number> c : List.of(colReqActualYen, colReqPlanYen)) {
            c.setCellFactory(col -> requestYenCell(false));
            c.setStyle("-fx-alignment: CENTER-RIGHT;");
        }
        colReqDiffYen.setCellFactory(col -> requestYenCell(true));
        colReqDiffYen.setStyle("-fx-alignment: CENTER-RIGHT;");
        requestTable.setPlaceholder(new Label("期間内の依頼がありません"));
        requestTable.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_ALL_COLUMNS);
        Tooltip.install(
                requestTable,
                new Tooltip(
                        "期間内の実績・予定を依頼NOごとに集約。AH単価欠落は —（当該依頼の円は 0）。"));
    }

    private TableCell<DayPoint, Number> yenCell(boolean signed) {
        return new TableCell<>() {
            @Override
            protected void updateItem(Number item, boolean empty) {
                super.updateItem(item, empty);
                pseudoClassStateChanged(PC_GOOD, false);
                pseudoClassStateChanged(PC_BAD, false);
                if (empty || item == null) {
                    setText(null);
                    return;
                }
                setText(formatYenNumber(item.doubleValue(), signed));
                if (signed) {
                    double v = item.doubleValue();
                    pseudoClassStateChanged(PC_GOOD, v > 0.5);
                    pseudoClassStateChanged(PC_BAD, v < -0.5);
                }
            }
        };
    }

    private TableCell<RequestPoint, Number> requestYenCell(boolean signed) {
        return new TableCell<>() {
            @Override
            protected void updateItem(Number item, boolean empty) {
                super.updateItem(item, empty);
                pseudoClassStateChanged(PC_GOOD, false);
                pseudoClassStateChanged(PC_BAD, false);
                if (empty || item == null) {
                    setText(null);
                    return;
                }
                setText(formatYenNumber(item.doubleValue(), signed));
                if (signed) {
                    double v = item.doubleValue();
                    pseudoClassStateChanged(PC_GOOD, v > 0.5);
                    pseudoClassStateChanged(PC_BAD, v < -0.5);
                }
            }
        };
    }

    private TableCell<RequestPoint, Number> metersCell() {
        return new TableCell<>() {
            @Override
            protected void updateItem(Number item, boolean empty) {
                super.updateItem(item, empty);
                if (empty || item == null) {
                    setText(null);
                    return;
                }
                NumberFormat nf = NumberFormat.getNumberInstance(Locale.JAPAN);
                nf.setMaximumFractionDigits(1);
                nf.setMinimumFractionDigits(0);
                setText(nf.format(item.doubleValue()));
            }
        };
    }

    private static String formatYenNumber(double v, boolean signed) {
        NumberFormat nf = NumberFormat.getIntegerInstance(Locale.JAPAN);
        if (signed) {
            String body = nf.format(Math.rint(Math.abs(v)));
            if (Math.abs(v) < 0.5) {
                return "±0";
            }
            return (v > 0 ? "+" : "−") + body;
        }
        return nf.format(Math.rint(v));
    }

    /** 軸目盛（加工量の formatM と同型。円もカンマ区切り整数）。 */
    private static String formatAxisTick(double v) {
        if (Math.abs(v) < 0.5) {
            return "0";
        }
        return String.format(Locale.JAPAN, "%,.0f", v);
    }

    private static String formatDayWithWeekday(LocalDate d) {
        return d.format(DateTimeFormatter.ofPattern("M/d"))
                + "("
                + d.getDayOfWeek().getDisplayName(TextStyle.SHORT, Locale.JAPAN)
                + ")";
    }

    private void initCumOverlayPaths() {
        styleCumPath(actualCumPath, "pm-trend-path-actual-cum", Color.web("#1e3a8a"), 2.5, false);
        styleCumPath(planCumPath, "pm-trend-path-fee-plan-cum", Color.web("#0f766e"), 2.5, true);
        markerPane.getChildren().add(0, planCumPath);
        markerPane.getChildren().add(0, actualCumPath);
    }

    private void initTodayMarker() {
        todayLine.getStyleClass().add("pm-processing-trend-today-line");
        todayLine.setManaged(false);
        todayMarkerLabel.getStyleClass().add("pm-processing-trend-today-label");
        todayMarkerLabel.setManaged(false);
        todayLine.setVisible(false);
        todayMarkerLabel.setVisible(false);
        markerPane.getChildren().addAll(todayLine, todayMarkerLabel);
    }

    private static void styleCumPath(
            Path path, String styleClass, Color stroke, double strokeWidth, boolean dashed) {
        path.setManaged(false);
        path.setMouseTransparent(true);
        path.setFill(null);
        path.setStroke(stroke);
        path.setStrokeWidth(strokeWidth);
        path.getStyleClass().setAll(styleClass);
        if (dashed) {
            path.getStrokeDashArray().setAll(6.0, 4.0);
        }
    }

    private void initLegend() {
        if (legendBox == null) {
            return;
        }
        legendBox.getChildren().setAll(
                legendItem("実績円", "#2563eb", false),
                legendItem("予定円", "#64748b", false),
                legendItem("実績累計", "#1e3a8a", false),
                legendItem("予定累計", "#0f766e", true),
                legendItem("今日", "#0f766e", true));
    }

    private static HBox legendItem(String text, String color, boolean dashed) {
        Region swatch = new Region();
        swatch.setMinSize(14, dashed ? 3 : 10);
        swatch.setPrefSize(14, dashed ? 3 : 10);
        if (dashed) {
            swatch.setStyle(
                    "-fx-border-color: "
                            + color
                            + " transparent transparent transparent; -fx-border-width: 2 0 0 0;");
        } else {
            swatch.setStyle("-fx-background-color: " + color + ";");
        }
        Label lab = new Label(text);
        lab.getStyleClass().add("pm-legend-label");
        HBox box = new HBox(6, swatch, lab);
        box.setAlignment(javafx.geometry.Pos.CENTER_LEFT);
        return box;
    }

    public void bindShell(MainShellController shell) {
        this.shell = shell;
    }

    public void onMainShellTabSelected() {
        tabActive = true;
        if (shell != null && shell.isStartupTabBackgroundLoadActive()) {
            return;
        }
        if (cachedSources == null) {
            reloadFromSources();
        } else {
            scheduleRecompute();
        }
    }

    public void onMainShellTabDeselected() {
        tabActive = false;
    }

    @FXML
    private void onReloadAction() {
        reloadFromSources();
    }

    @FXML
    private void onThisMonthAction() {
        suppressFilterEvents = true;
        applyPreset(PeriodPreset.THIS_MONTH);
        periodPresetCombo.getSelectionModel().select(PeriodPreset.THIS_MONTH);
        suppressFilterEvents = false;
        scheduleRecompute();
    }

    private void onDateManual() {
        if (suppressFilterEvents) {
            return;
        }
        suppressFilterEvents = true;
        periodPresetCombo.getSelectionModel().select(PeriodPreset.CUSTOM);
        suppressFilterEvents = false;
        scheduleRecompute();
    }

    private void applyPreset(PeriodPreset p) {
        LocalDate today = LocalDate.now();
        LocalDate from;
        LocalDate to;
        switch (p) {
            case LAST_MONTH -> {
                YearMonth ym = YearMonth.from(today).minusMonths(1);
                from = ym.atDay(1);
                to = ym.atEndOfMonth();
            }
            case PAST_3_MONTHS -> {
                to = today;
                from = today.minusMonths(3).plusDays(1);
            }
            default -> {
                YearMonth ym = YearMonth.from(today);
                from = ym.atDay(1);
                to = ym.atEndOfMonth();
            }
        }
        fromDatePicker.setValue(from);
        toDatePicker.setValue(to);
    }

    private void scheduleRecompute() {
        if (suppressFilterEvents || cachedSources == null) {
            return;
        }
        recomputeNow();
    }

    private void reloadFromSources() {
        if (shell == null || reloadInFlight) {
            return;
        }
        reloadInFlight = true;
        setLoading(true, "ソース読込中…");
        final Map<String, String> ui = shell.snapshotUiEnv();
        final SourceFingerprint previous = loadedFingerprint;
        final boolean haveCache = cachedSources != null;
        Task<ReloadBundle> task =
                new Task<>() {
                    @Override
                    protected ReloadBundle call() throws Exception {
                        ReloadDecision decision =
                                EquipmentStatusDashboardSourceLoader.loadIfChanged(ui, previous, haveCache);
                        LoadedSources sources =
                                decision != null && decision.sources() != null
                                        ? decision.sources()
                                        : (haveCache ? cachedSources : null);
                        Map<String, Double> rates = Map.of();
                        java.nio.file.Path juchu =
                                AppPaths.resolveRequestFormJuchuFile(ui).orElse(null);
                        String juchuNote;
                        if (juchu != null && Files.isRegularFile(juchu)) {
                            rates = JuchuProcessingFeeRateLoader.loadRates(juchu);
                            juchuNote = juchu.getFileName().toString() + " AH=" + rates.size() + "件";
                        } else {
                            juchuNote = "受注ファイル未設定または不在";
                        }
                        return new ReloadBundle(decision, sources, rates, juchuNote);
                    }
                };
        task.setOnSucceeded(
                e -> {
                    reloadInFlight = false;
                    setLoading(false, "");
                    ReloadBundle b = task.getValue();
                    if (b.decision() != null && !b.decision().sourcesUnchanged() && b.sources() != null) {
                        cachedSources = b.sources();
                        loadedFingerprint = b.decision().fingerprint();
                    } else if (b.sources() != null) {
                        cachedSources = b.sources();
                    }
                    feeRates = b.rates() != null ? b.rates() : Map.of();
                    sourceSummaryLabel.setText(
                            "単価: "
                                    + b.juchuNote()
                                    + " ／ 単位: 円（AH×工程延べ m）。依頼の生産金額ではありません。");
                    hideNotice();
                    recomputeNow();
                });
        task.setOnFailed(
                e -> {
                    reloadInFlight = false;
                    setLoading(false, "");
                    Throwable ex = task.getException();
                    showNotice("読込失敗: " + (ex != null ? ex.getMessage() : "不明"));
                });
        Thread th = new Thread(task, "processing-fee-trend-reload");
        th.setDaemon(true);
        th.start();
    }

    private void recomputeNow() {
        if (cachedSources == null) {
            renderEmpty();
            return;
        }
        LocalDate from = fromDatePicker.getValue();
        LocalDate to = toDatePicker.getValue();
        if (from == null || to == null) {
            return;
        }
        if (ChronoUnit.DAYS.between(from, to) > 400) {
            showNotice("期間が長すぎます（最大400日）。");
            return;
        }
        final long seq = computeSeq.incrementAndGet();
        final LoadedSources src = cachedSources;
        final Map<String, Double> rates = feeRates;
        final ActualSource actSrc = actualSourceCombo.getValue();
        final PlanSource planSrc = planSourceCombo.getValue();
        setLoading(true, "集計中…");
        Task<Result> task =
                new Task<>() {
                    @Override
                    protected Result call() {
                        Filter filter = new Filter(from, to, actSrc, planSrc, null, null, 7);
                        List<QuantityLine> actual =
                                ProcessingFeeTrendQuantityExtractor.extractActual(
                                        src.dailyReportActuals(), src.actuals(), filter);
                        List<QuantityLine> plan =
                                ProcessingFeeTrendQuantityExtractor.extractPlan(
                                        src.aladdin(), src.dispatch(), filter, LocalDate.now());
                        return ProcessingFeeTrendAggregator.aggregate(
                                actual, plan, rates, from, to, LocalDate.now());
                    }
                };
        task.setOnSucceeded(
                e -> {
                    if (seq != computeSeq.get()) {
                        return;
                    }
                    setLoading(false, "");
                    Result r = task.getValue();
                    currentResult = r;
                    render(r);
                    if (r.missingRateLines() > 0) {
                        showNotice(
                                "AH 単価が無い行が "
                                        + r.missingRateLines()
                                        + " 件あります（当該行は 0 円）。");
                    } else {
                        hideNotice();
                    }
                });
        task.setOnFailed(
                e -> {
                    if (seq != computeSeq.get()) {
                        return;
                    }
                    setLoading(false, "");
                    Throwable ex = task.getException();
                    showNotice("集計失敗: " + (ex != null ? ex.getMessage() : "不明"));
                });
        Thread th = new Thread(task, "processing-fee-trend-compute");
        th.setDaemon(true);
        th.start();
    }

    private void render(Result r) {
        List<DayPoint> days = r.days();
        int n = days.size();
        boolean singleYear = n > 0 && days.get(0).date().getYear() == days.get(n - 1).date().getYear();
        boolean singleMonth =
                singleYear && days.get(0).date().getMonth() == days.get(n - 1).date().getMonth();
        List<String> labels = new ArrayList<>(n);
        List<LocalDate> dates = new ArrayList<>(n);
        for (DayPoint d : days) {
            labels.add(ProcessingTrendChartSupport.categoryLabel(d.date(), singleMonth, singleYear));
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

        var cats = FXCollections.observableArrayList(labels);
        dailyXAxis.setCategories(cats);
        if (cumulativeXAxis != null) {
            cumulativeXAxis.setCategories(FXCollections.observableArrayList(labels));
        }
        dailyChart.setCategoryGap(ProcessingTrendChartSupport.categoryGapFor(n));
        dailyChart.setBarGap(ProcessingTrendChartSupport.barGapFor(n));

        List<XYChart.Data<String, Number>> act = new ArrayList<>(n);
        List<XYChart.Data<String, Number>> plan = new ArrayList<>(n);
        List<XYChart.Data<String, Number>> actCum = new ArrayList<>(n);
        List<XYChart.Data<String, Number>> planCum = new ArrayList<>(n);
        double dailyMax = 0;
        double cumMax = 0;
        LocalDate today = r.today();
        for (int i = 0; i < n; i++) {
            DayPoint d = days.get(i);
            String cat = labels.get(i);
            act.add(new XYChart.Data<>(cat, d.actualYen()));
            plan.add(new XYChart.Data<>(cat, d.planYen()));
            dailyMax = Math.max(dailyMax, Math.max(d.actualYen(), d.planYen()));
            cumMax = Math.max(cumMax, Math.max(d.actualCumYen(), d.planCumYen()));
            // 実績累計は当日まで（未来へ水平延長しない）
            if (!d.date().isAfter(today)) {
                actCum.add(new XYChart.Data<>(cat, d.actualCumYen()));
            }
            planCum.add(new XYChart.Data<>(cat, d.planCumYen()));
        }
        // 第一軸（左）= 日次棒 / 第二軸（右）= 累計折線（LineChart 系列で右軸を確実に立てる）
        applyNiceRange(dailyYAxis, dailyMax);
        applyNiceRange(cumulativeYAxis, cumMax);
        actualSeries.getData().setAll(act);
        planSeries.getData().setAll(plan);
        actualCumSeries.getData().setAll(actCum);
        planCumSeries.getData().setAll(planCum);
        // Path は使わず LineChart の右軸スケールで描く
        overlayActualCum = OverlayPolyline.EMPTY;
        overlayPlanCum = OverlayPolyline.EMPTY;
        hideCumPaths();
        styleCumSeriesLines();
        applyChartWhiteBackground();

        NumberFormat nf = NumberFormat.getNumberInstance(Locale.JAPAN);
        nf.setMaximumFractionDigits(0);
        kpiActualYen.setText(nf.format(Math.rint(r.actualTotalYen())) + " 円");
        kpiPlanYen.setText(nf.format(Math.rint(r.planTotalYen())) + " 円");

        if (detailTable != null) {
            detailTable.getItems().setAll(days);
            detailTable.refresh();
        }
        if (requestTable != null) {
            requestTable.getItems().setAll(r.requests());
            requestTable.refresh();
        }
        if (requestPane != null) {
            requestPane.setText("依頼NO別 加工賃（" + r.requests().size() + " 件）");
        }

        settleOverlayLayout();
    }

    private static void applyNiceRange(ProcessingTrendNumberAxis axis, double max) {
        NiceRange nr = ProcessingTrendChartSupport.niceRange(max);
        axis.setLowerBound(0);
        axis.setUpperBound(nr.upperBound());
        axis.setTickUnit(nr.tickUnit());
    }

    private void syncChartPadding() {
        // 加工量 COMBO と同型: 実測幅のみ（minWidth を余白に食い込ませない）
        double left = dailyYAxis.getWidth();
        double right = cumulativeYAxis.getWidth();
        dailyChart.setPadding(new Insets(CHART_TOP_PADDING, right, 0, 0));
        cumulativeChart.setPadding(new Insets(CHART_TOP_PADDING, 0, 0, left));
    }

    private void settleOverlayLayout() {
        syncChartPadding();
        applyChartWhiteBackground();
        requestOverlayLayout();
        Platform.runLater(
                () -> {
                    syncChartPadding();
                    applyChartWhiteBackground();
                    styleCumSeriesLines();
                    requestOverlayLayout();
                    Platform.runLater(
                            () -> {
                                syncChartPadding();
                                applyChartWhiteBackground();
                                requestOverlayLayout();
                            });
                });
    }

    /** CSS がテーマに負ける場合でもプロットを白地にする（加工量チャートと同見た目）。 */
    private void applyChartWhiteBackground() {
        if (chartStack != null) {
            chartStack.setBackground(CHART_STACK_WHITE);
        }
        paintChartChrome(dailyChart, PLOT_WHITE);
        // 上層は透過（下の棒を隠さない）
        paintChartChrome(cumulativeChart, PLOT_TRANSPARENT);
    }

    private static void paintChartChrome(XYChart<?, ?> chart, Background plotBg) {
        if (chart == null) {
            return;
        }
        chart.setStyle("-fx-background-color: transparent;");
        Node plot = chart.lookup(".chart-plot-background");
        if (plot instanceof Region region) {
            region.setBackground(plotBg);
            region.setStyle(
                    plotBg == PLOT_WHITE
                            ? "-fx-background-color: white;"
                            : "-fx-background-color: transparent;");
        }
        Node content = chart.lookup(".chart-content");
        if (content instanceof Region region) {
            region.setStyle("-fx-background-color: transparent;");
        }
    }

    private void styleCumSeriesLines() {
        styleSeriesLine(actualCumSeries, "#1e3a8a", false);
        styleSeriesLine(planCumSeries, "#0f766e", true);
    }

    private static void styleSeriesLine(
            XYChart.Series<String, Number> series, String stroke, boolean dashed) {
        if (series == null) {
            return;
        }
        Node node = series.getNode();
        if (node == null) {
            return;
        }
        String dash = dashed ? "; -fx-stroke-dash-array: 6 4" : "";
        String css = "-fx-stroke: " + stroke + "; -fx-stroke-width: 2.5px" + dash + ";";
        node.setStyle(css);
        for (Node child : node.lookupAll(".chart-series-line")) {
            child.setStyle(css);
        }
    }

    private void requestOverlayLayout() {
        if (overlayLayoutScheduled) {
            return;
        }
        overlayLayoutScheduled = true;
        Platform.runLater(
                () -> {
                    overlayLayoutScheduled = false;
                    layoutOverlay();
                });
    }

    private void layoutOverlay() {
        Result r = currentResult;
        if (r == null || r.days().isEmpty() || dailyXAxis.getScene() == null) {
            hideXAxisLabels();
            hideCumPaths();
            todayLine.setVisible(false);
            todayMarkerLabel.setVisible(false);
            return;
        }
        Node plotBg = dailyChart.lookup(".chart-plot-background");
        if (plotBg == null || markerPane.getScene() == null) {
            hideXAxisLabels();
            hideCumPaths();
            todayLine.setVisible(false);
            todayMarkerLabel.setVisible(false);
            return;
        }
        layoutXAxisLabels(dailyXAxis);
        layoutTodayMarker(r, dailyXAxis, plotBg);
        // 累計線は cumulativeChart 系列（右軸）。Path は使わない。
        hideCumPaths();
    }

    private void hideCumPaths() {
        actualCumPath.getElements().clear();
        planCumPath.getElements().clear();
        actualCumPath.setVisible(false);
        planCumPath.setVisible(false);
    }

    private void layoutCumPaths() {
        layoutOneCumPath(actualCumPath, overlayActualCum);
        layoutOneCumPath(planCumPath, overlayPlanCum);
    }

    private void layoutOneCumPath(Path path, OverlayPolyline poly) {
        path.getElements().clear();
        if (poly == null || poly.isEmpty() || dailyXAxis.getScene() == null) {
            path.setVisible(false);
            return;
        }
        if (cumulativeYAxis.getScene() == null || markerPane.getScene() == null) {
            path.setVisible(false);
            return;
        }
        boolean started = false;
        for (int i = 0; i < poly.categories().size(); i++) {
            double ax = dailyXAxis.getDisplayPosition(poly.categories().get(i));
            double ay = cumulativeYAxis.getDisplayPosition(poly.values().get(i));
            if (Double.isNaN(ax) || Double.isNaN(ay)) {
                continue;
            }
            Point2D xs = dailyXAxis.localToScene(ax, 0);
            Point2D ys = cumulativeYAxis.localToScene(0, ay);
            Point2D p = markerPane.sceneToLocal(xs.getX(), ys.getY());
            if (Double.isNaN(p.getX()) || Double.isNaN(p.getY())) {
                continue;
            }
            if (!started) {
                path.getElements().add(new MoveTo(p.getX(), p.getY()));
                started = true;
            } else {
                path.getElements().add(new LineTo(p.getX(), p.getY()));
            }
        }
        path.setVisible(started);
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
            int idx = currentCategoryLabels.size() - 1;
            x =
                    xAxis.getDisplayPosition(currentCategoryLabels.get(idx))
                            + xAxis.getCategorySpacing() / 2.0;
        } else {
            int idx = (int) ChronoUnit.DAYS.between(first, today);
            if (idx < 0 || idx >= currentCategoryLabels.size()) {
                todayLine.setVisible(false);
                todayMarkerLabel.setVisible(false);
                return;
            }
            x =
                    xAxis.getDisplayPosition(currentCategoryLabels.get(idx))
                            - xAxis.getCategorySpacing() / 2.0;
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
        todayMarkerLabel.setText("今日");
        todayMarkerLabel.applyCss();
        todayMarkerLabel.autosize();
        double labelW = todayMarkerLabel.getWidth();
        double lx = px + 4;
        if (lx + labelW > plotBounds.getMaxX() - 2) {
            lx = px - labelW - 4;
        }
        todayMarkerLabel.relocate(Math.round(lx), Math.round(plotBounds.getMinY() + 2));
        todayMarkerLabel.setVisible(true);
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
            String key = currentCategoryLabels.get(i);
            l.setText(key != null && key.matches("0[1-9]") ? key.substring(1) : key);
            LocalDate d = currentDates.get(i);
            l.pseudoClassStateChanged(PC_SAT, d.getDayOfWeek() == DayOfWeek.SATURDAY);
            l.pseudoClassStateChanged(PC_SUN, d.getDayOfWeek() == DayOfWeek.SUNDAY);
            l.applyCss();
            l.autosize();
            double w = l.getWidth();
            double lx = Math.max(axisMinX, Math.min(p.getX() - w / 2.0, axisMaxX - w));
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

    private void hideXAxisLabels() {
        for (Label l : xAxisLabelPool) {
            l.setVisible(false);
        }
    }

    private void renderEmpty() {
        currentResult = null;
        actualSeries.getData().clear();
        planSeries.getData().clear();
        actualCumSeries.getData().clear();
        planCumSeries.getData().clear();
        overlayActualCum = OverlayPolyline.EMPTY;
        overlayPlanCum = OverlayPolyline.EMPTY;
        hideCumPaths();
        hideXAxisLabels();
        todayLine.setVisible(false);
        todayMarkerLabel.setVisible(false);
        kpiActualYen.setText("—");
        kpiPlanYen.setText("—");
        if (detailTable != null) {
            detailTable.getItems().clear();
        }
        if (requestTable != null) {
            requestTable.getItems().clear();
        }
        if (requestPane != null) {
            requestPane.setText("依頼NO別 加工賃");
        }
    }

    private void setLoading(boolean on, String text) {
        loadingChip.setVisible(on);
        loadingChip.setManaged(on);
        if (text != null && !text.isEmpty()) {
            loadingStatusLabel.setText(text);
        }
        if (reloadButton != null) {
            reloadButton.setDisable(on);
        }
    }

    private void showNotice(String msg) {
        noticeLabel.setText(msg);
        noticeBanner.setVisible(true);
        noticeBanner.setManaged(true);
    }

    private void hideNotice() {
        noticeBanner.setVisible(false);
        noticeBanner.setManaged(false);
        noticeLabel.setText("");
    }

    private record ReloadBundle(
            ReloadDecision decision,
            LoadedSources sources,
            Map<String, Double> rates,
            String juchuNote) {}
}
