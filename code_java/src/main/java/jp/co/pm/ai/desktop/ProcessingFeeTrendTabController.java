package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.text.NumberFormat;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.concurrent.atomic.AtomicLong;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.geometry.Point2D;
import javafx.scene.chart.BarChart;
import javafx.scene.chart.CategoryAxis;
import javafx.scene.chart.XYChart;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.DatePicker;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Pane;
import javafx.scene.layout.StackPane;
import javafx.scene.paint.Color;
import javafx.scene.shape.LineTo;
import javafx.scene.shape.MoveTo;
import javafx.scene.shape.Path;

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
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.Result;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendQuantityExtractor;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.ActualSource;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.Filter;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.PlanSource;

/**
 * 「加工賃トレンド」タブ: AH×m の実績円・予定円・累計折れ線（独立グラフ）。
 */
public class ProcessingFeeTrendTabController {

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
    @FXML private StackPane chartStack;
    @FXML private BarChart<String, Number> dailyChart;
    @FXML private CategoryAxis dailyXAxis;
    @FXML private ProcessingTrendNumberAxis dailyYAxis;
    @FXML private Pane markerPane;
    @FXML private Label sourceSummaryLabel;

    private MainShellController shell;
    private boolean suppressFilterEvents;
    private boolean tabActive;
    private boolean reloadInFlight;
    private LoadedSources cachedSources;
    private SourceFingerprint loadedFingerprint;
    private Map<String, Double> feeRates = Map.of();
    private Result currentResult;
    private final AtomicLong computeSeq = new AtomicLong();

    private final XYChart.Series<String, Number> actualSeries = new XYChart.Series<>();
    private final XYChart.Series<String, Number> planSeries = new XYChart.Series<>();
    private final Path actualCumPath = new Path();
    private final Path planCumPath = new Path();

    @FXML
    private void initialize() {
        actualSeries.setName("実績円");
        planSeries.setName("予定円");
        dailyChart.getData().setAll(actualSeries, planSeries);
        dailyChart.setAnimated(false);

        stylePath(actualCumPath, Color.web("#1e3a8a"), 2.5, false);
        stylePath(planCumPath, Color.web("#0f766e"), 2.5, true);
        markerPane.getChildren().addAll(actualCumPath, planCumPath);

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

        dailyChart.layoutBoundsProperty().addListener((o, a, n) -> layoutCumPaths());
        dailyXAxis.layoutBoundsProperty().addListener((o, a, n) -> layoutCumPaths());
        dailyYAxis.widthProperty().addListener((o, a, n) -> layoutCumPaths());

        sourceSummaryLabel.setText("受注 AH × 加工日 m。工程延べのため依頼生産量とは一致しません。");
        renderEmpty();
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

    public void preloadInBackground(java.util.function.Consumer<Boolean> onComplete) {
        if (shell == null) {
            if (onComplete != null) {
                Platform.runLater(() -> onComplete.accept(false));
            }
            return;
        }
        reloadFromSources(onComplete);
    }

    @FXML
    private void onReloadAction() {
        reloadFromSources(null);
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
        reloadFromSources(null);
    }

    private void reloadFromSources(java.util.function.Consumer<Boolean> onComplete) {
        if (shell == null || reloadInFlight) {
            if (onComplete != null) {
                Platform.runLater(() -> onComplete.accept(false));
            }
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
                        String juchuNote = "";
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
                    if (onComplete != null) {
                        onComplete.accept(cachedSources != null);
                    }
                });
        task.setOnFailed(
                e -> {
                    reloadInFlight = false;
                    setLoading(false, "");
                    Throwable ex = task.getException();
                    showNotice("読込失敗: " + (ex != null ? ex.getMessage() : "不明"));
                    if (onComplete != null) {
                        onComplete.accept(false);
                    }
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
                        Filter filter =
                                new Filter(from, to, actSrc, planSrc, null, null, 7);
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
        for (DayPoint d : days) {
            labels.add(ProcessingTrendChartSupport.categoryLabel(d.date(), singleMonth, singleYear));
        }
        dailyXAxis.setCategories(FXCollections.observableArrayList(labels));
        dailyChart.setCategoryGap(ProcessingTrendChartSupport.categoryGapFor(n));

        List<XYChart.Data<String, Number>> act = new ArrayList<>(n);
        List<XYChart.Data<String, Number>> plan = new ArrayList<>(n);
        double max = 0;
        for (int i = 0; i < n; i++) {
            DayPoint d = days.get(i);
            String cat = labels.get(i);
            act.add(new XYChart.Data<>(cat, d.actualYen()));
            plan.add(new XYChart.Data<>(cat, d.planYen()));
            max = Math.max(max, Math.max(d.actualYen(), Math.max(d.planYen(), d.planCumYen())));
            max = Math.max(max, d.actualCumYen());
        }
        NiceRange nr = ProcessingTrendChartSupport.niceRange(max);
        dailyYAxis.setLowerBound(0);
        dailyYAxis.setUpperBound(nr.upperBound());
        dailyYAxis.setTickUnit(nr.tickUnit());
        actualSeries.getData().setAll(act);
        planSeries.getData().setAll(plan);

        NumberFormat nf = NumberFormat.getNumberInstance(Locale.JAPAN);
        nf.setMaximumFractionDigits(0);
        kpiActualYen.setText(nf.format(Math.rint(r.actualTotalYen())) + " 円");
        kpiPlanYen.setText(nf.format(Math.rint(r.planTotalYen())) + " 円");
        chartTitleLabel.setText("加工賃トレンド（円）");

        Platform.runLater(
                () -> {
                    layoutCumPaths();
                    Platform.runLater(this::layoutCumPaths);
                });
    }

    private void layoutCumPaths() {
        Result r = currentResult;
        actualCumPath.getElements().clear();
        planCumPath.getElements().clear();
        if (r == null || r.days().isEmpty() || dailyXAxis.getScene() == null) {
            actualCumPath.setVisible(false);
            planCumPath.setVisible(false);
            return;
        }
        List<String> cats = dailyXAxis.getCategories();
        if (cats == null || cats.size() != r.days().size()) {
            return;
        }
        buildPath(actualCumPath, cats, r.days(), true);
        buildPath(planCumPath, cats, r.days(), false);
    }

    private void buildPath(Path path, List<String> cats, List<DayPoint> days, boolean actual) {
        boolean started = false;
        LocalDate today = currentResult != null ? currentResult.today() : LocalDate.now();
        for (int i = 0; i < days.size(); i++) {
            DayPoint d = days.get(i);
            if (actual && d.date().isAfter(today)) {
                break;
            }
            double yVal = actual ? d.actualCumYen() : d.planCumYen();
            double ax = dailyXAxis.getDisplayPosition(cats.get(i));
            double ay = dailyYAxis.getDisplayPosition(yVal);
            if (Double.isNaN(ax) || Double.isNaN(ay)) {
                continue;
            }
            Point2D xs = dailyXAxis.localToScene(ax, 0);
            Point2D ys = dailyYAxis.localToScene(0, ay);
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

    private void renderEmpty() {
        currentResult = null;
        actualSeries.getData().clear();
        planSeries.getData().clear();
        actualCumPath.getElements().clear();
        planCumPath.getElements().clear();
        kpiActualYen.setText("—");
        kpiPlanYen.setText("—");
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

    private static void stylePath(Path path, Color stroke, double width, boolean dashed) {
        path.setManaged(false);
        path.setMouseTransparent(true);
        path.setFill(null);
        path.setStroke(stroke);
        path.setStrokeWidth(width);
        if (dashed) {
            path.getStrokeDashArray().setAll(6.0, 4.0);
        }
    }

    private record ReloadBundle(
            ReloadDecision decision,
            LoadedSources sources,
            Map<String, Double> rates,
            String juchuNote) {}
}
