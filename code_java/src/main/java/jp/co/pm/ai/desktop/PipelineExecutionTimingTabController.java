package jp.co.pm.ai.desktop;

import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.EnumMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import javafx.application.Platform;
import javafx.beans.property.ReadOnlyStringWrapper;
import javafx.collections.FXCollections;
import javafx.fxml.FXML;
import javafx.scene.chart.BarChart;
import javafx.scene.chart.LineChart;
import javafx.scene.chart.XYChart;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.ToggleButton;
import javafx.scene.control.ToggleGroup;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.PipelineExecutionTimingHistoryStore.Stats;

/**
 * パイプライン実行時間のトレンド・分布（ヒストグラム）を表示するメインシェルタブ。
 */
public final class PipelineExecutionTimingTabController {

    private static final DateTimeFormatter FINISHED_AT =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss").withZone(ZoneId.systemDefault());

    /** トレンドで全種別を重ね表示するときのコンボ選択肢。 */
    private static final String KIND_FILTER_ALL = "すべて（重ね表示）";

    private MainShellController shell;

    private final Runnable historyChangeListener = () -> Platform.runLater(this::refreshFromStore);

    @FXML
    private Label storageDetailLabel;

    @FXML
    private Label hintLabel;

    @FXML
    private ComboBox<String> kindFilterCombo;

    @FXML
    private Spinner<Integer> sampleLimitSpinner;

    @FXML
    private Button refreshChartsButton;

    @FXML
    private Button clearHistoryButton;

    @FXML
    private TableView<StatsRow> statsTable;

    @FXML
    private TableColumn<StatsRow, String> statsKindColumn;

    @FXML
    private TableColumn<StatsRow, String> statsCountColumn;

    @FXML
    private TableColumn<StatsRow, String> statsAvgColumn;

    @FXML
    private TableColumn<StatsRow, String> statsMedianColumn;

    @FXML
    private TableColumn<StatsRow, String> statsMinColumn;

    @FXML
    private TableColumn<StatsRow, String> statsMaxColumn;

    @FXML
    private TableColumn<StatsRow, String> statsStdDevColumn;

    @FXML
    private ToggleGroup chartViewToggleGroup;

    @FXML
    private ToggleButton trendViewToggle;

    @FXML
    private ToggleButton distributionViewToggle;

    @FXML
    private LineChart<Number, Number> trendChart;

    @FXML
    private BarChart<String, Number> distributionChart;

    @FXML
    private TableView<HistoryRow> historyTable;

    @FXML
    private TableColumn<HistoryRow, String> historyTimeColumn;

    @FXML
    private TableColumn<HistoryRow, String> historyKindColumn;

    @FXML
    private TableColumn<HistoryRow, String> historyDurationColumn;

    @FXML
    private TableColumn<HistoryRow, String> historyWriterColumn;

    private final Map<PipelineExecutionTimingKind, XYChart.Series<Number, Number>> trendSeriesByKind =
            new EnumMap<>(PipelineExecutionTimingKind.class);

    @FXML
    private void initialize() {
        rebuildKindChoices();
        kindFilterCombo
                .getSelectionModel()
                .selectedItemProperty()
                .addListener((o, a, b) -> refreshFromStore());

        sampleLimitSpinner.setValueFactory(
                new SpinnerValueFactory.IntegerSpinnerValueFactory(10, 500, 100, 10));
        sampleLimitSpinner.valueProperty().addListener((o, a, b) -> refreshFromStore());

        bindStatsColumns();
        statsTable.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_FLEX_LAST_COLUMN);
        historyTable.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_FLEX_LAST_COLUMN);

        historyTimeColumn.setCellValueFactory(
                cd -> new ReadOnlyStringWrapper(cd.getValue().finishedAtText()));
        historyKindColumn.setCellValueFactory(
                cd -> new ReadOnlyStringWrapper(cd.getValue().kindLabel()));
        historyDurationColumn.setCellValueFactory(
                cd -> new ReadOnlyStringWrapper(cd.getValue().durationText()));
        historyWriterColumn.setCellValueFactory(
                cd -> new ReadOnlyStringWrapper(cd.getValue().writerText()));

        chartViewToggleGroup
                .selectedToggleProperty()
                .addListener((o, a, b) -> applyChartViewVisibility());

        for (PipelineExecutionTimingKind kind : PipelineExecutionTimingKind.values()) {
            XYChart.Series<Number, Number> series = new XYChart.Series<>();
            series.setName(kind.label());
            trendSeriesByKind.put(kind, series);
            if (isKindVisible(kind)) {
                trendChart.getData().add(series);
            }
        }
        applyChartViewVisibility();
        if (hintLabel != null) {
            hintLabel.setText(
                    "段階1・2.0～2.1・配台試行・サマリ Excel・納期管理ビューの所要時間を記録・表示します。");
        }
    }

    private void rebuildKindChoices() {
        String selected = kindFilterCombo.getSelectionModel().getSelectedItem();
        List<String> kindChoices = new ArrayList<>();
        kindChoices.add(KIND_FILTER_ALL);
        for (PipelineExecutionTimingKind kind : PipelineExecutionTimingKind.values()) {
            if (isKindVisible(kind)) {
                kindChoices.add(kind.label());
            }
        }
        kindFilterCombo.setItems(FXCollections.observableArrayList(kindChoices));
        if (selected == null || !kindChoices.contains(selected)) {
            kindFilterCombo.getSelectionModel().selectFirst();
        } else {
            kindFilterCombo.getSelectionModel().select(selected);
        }
    }

    private boolean isKindVisible(PipelineExecutionTimingKind kind) {
        return kind != null;
    }

    private void bindStatsColumns() {
        statsKindColumn.setCellValueFactory(cd -> new ReadOnlyStringWrapper(cd.getValue().kindLabel()));
        statsCountColumn.setCellValueFactory(cd -> new ReadOnlyStringWrapper(cd.getValue().countText()));
        statsAvgColumn.setCellValueFactory(cd -> new ReadOnlyStringWrapper(cd.getValue().avgText()));
        statsMedianColumn.setCellValueFactory(cd -> new ReadOnlyStringWrapper(cd.getValue().medianText()));
        statsMinColumn.setCellValueFactory(cd -> new ReadOnlyStringWrapper(cd.getValue().minText()));
        statsMaxColumn.setCellValueFactory(cd -> new ReadOnlyStringWrapper(cd.getValue().maxText()));
        statsStdDevColumn.setCellValueFactory(cd -> new ReadOnlyStringWrapper(cd.getValue().stdDevText()));
        statsCountColumn.setStyle("-fx-alignment: CENTER-RIGHT;");
        statsAvgColumn.setStyle("-fx-alignment: CENTER-RIGHT;");
        statsMedianColumn.setStyle("-fx-alignment: CENTER-RIGHT;");
        statsMinColumn.setStyle("-fx-alignment: CENTER-RIGHT;");
        statsMaxColumn.setStyle("-fx-alignment: CENTER-RIGHT;");
        statsStdDevColumn.setStyle("-fx-alignment: CENTER-RIGHT;");
    }

    private void applyChartViewVisibility() {
        boolean trend = trendViewToggle != null && trendViewToggle.isSelected();
        if (trendChart != null) {
            trendChart.setVisible(trend);
            trendChart.setManaged(trend);
        }
        if (distributionChart != null) {
            distributionChart.setVisible(!trend);
            distributionChart.setManaged(!trend);
        }
    }

    void bindShell(MainShellController shell) {
        if (this.shell != null) {
            this.shell.pipelineExecutionTimingHistory().removeChangeListener(historyChangeListener);
        }
        this.shell = shell;
        if (shell != null) {
            shell.pipelineExecutionTimingHistory().addChangeListener(historyChangeListener);
        }
        refreshFromStore();
    }

    @FXML
    private void onRefreshChartsAction() {
        refreshFromStore();
    }

    @FXML
    private void onClearHistoryAction() {
        if (shell == null) {
            return;
        }
        PipelineExecutionTimingKind selectedKind = selectedKindOrNull();
        String message =
                selectedKind == null
                        ? "すべての実行時間履歴を削除します。よろしいですか？"
                        : "「" + selectedKind.label() + "」の履歴のみ削除します。よろしいですか？";
        Alert alert = new Alert(Alert.AlertType.CONFIRMATION, message, ButtonType.OK, ButtonType.CANCEL);
        alert.setTitle("履歴クリア");
        alert.setHeaderText(null);
        shell.prepareDialogForMainTheme(alert);
        alert.showAndWait().ifPresent(btn -> {
            if (btn != ButtonType.OK) {
                return;
            }
            if (selectedKind == null) {
                shell.pipelineExecutionTimingHistory().clearAll();
            } else {
                shell.pipelineExecutionTimingHistory().clearKind(selectedKind);
            }
        });
    }

    void refreshFromStore() {
        if (shell == null) {
            return;
        }
        PipelineExecutionTimingHistoryStore store = shell.pipelineExecutionTimingHistory();
        refreshStorageDetail(store);
        int limit = sampleLimitSpinner.getValue() != null ? sampleLimitSpinner.getValue() : 100;
        PipelineExecutionTimingKind selectedKind = selectedKindOrNull();

        List<PipelineExecutionTimingSample> tableSamples =
                selectedKind == null
                        ? store.recentSamples(null, limit)
                        : store.recentSamplesForKind(selectedKind, limit);
        tableSamples = tableSamples.stream().filter(s -> isKindVisible(s.kind())).toList();
        refreshHistoryTable(tableSamples);
        refreshTrendChart(store, selectedKind, limit);
        refreshDistributionChart(selectedKind, limit, tableSamples);
        refreshStatsTable(selectedKind, limit);
    }

    private void refreshStorageDetail(PipelineExecutionTimingHistoryStore store) {
        if (storageDetailLabel == null) {
            return;
        }
        String lockLine =
                store.isPersistLocked()
                        ? store.readPersistLockInfo()
                                .map(
                                        info ->
                                                "他端末が保存中: "
                                                        + info.displayHost()
                                                        + " / "
                                                        + info.displayIp())
                                .orElse("他端末が保存中")
                        : "ロック: なし";
        storageDetailLabel.setText(
                "履歴 JSON はサマリ Excel（"
                        + AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK
                        + "）と同一フォルダに保存されます。保存前に .save.lock で排他します。\n"
                        + lockLine
                        + "\n"
                        + store.storagePath());
    }

    private PipelineExecutionTimingKind selectedKindOrNull() {
        String selected = kindFilterCombo.getSelectionModel().getSelectedItem();
        if (selected == null || KIND_FILTER_ALL.equals(selected)) {
            return null;
        }
        for (PipelineExecutionTimingKind kind : PipelineExecutionTimingKind.values()) {
            if (kind.label().equals(selected)) {
                return kind;
            }
        }
        return null;
    }

    private void refreshHistoryTable(List<PipelineExecutionTimingSample> samples) {
        List<HistoryRow> rows = new ArrayList<>(samples.size());
        for (int i = samples.size() - 1; i >= 0; i--) {
            PipelineExecutionTimingSample s = samples.get(i);
            rows.add(
                    new HistoryRow(
                            FINISHED_AT.format(Instant.ofEpochMilli(s.finishedAtEpochMs())),
                            s.kind().label(),
                            MainRunTabController.formatPipelineExecutionDuration(s.durationMs()),
                            s.writerEndpointLabel()));
        }
        historyTable.setItems(FXCollections.observableArrayList(rows));
    }

    private void refreshTrendChart(
            PipelineExecutionTimingHistoryStore store,
            PipelineExecutionTimingKind selectedKind,
            int limit) {
        for (PipelineExecutionTimingKind kind : PipelineExecutionTimingKind.values()) {
            XYChart.Series<Number, Number> series = trendSeriesByKind.get(kind);
            series.getData().clear();
            if (!isKindVisible(kind)) {
                continue;
            }
            boolean show = selectedKind == null || selectedKind == kind;
            if (!show) {
                continue;
            }
            List<PipelineExecutionTimingSample> samples = store.recentSamplesForKind(kind, limit);
            int run = 1;
            for (PipelineExecutionTimingSample sample : samples) {
                double sec = sample.durationMs() / 1000.0;
                series.getData().add(new XYChart.Data<>(run, sec));
                run++;
            }
        }
    }

    private void refreshDistributionChart(
            PipelineExecutionTimingKind selectedKind,
            int limit,
            List<PipelineExecutionTimingSample> tableSamples) {
        distributionChart.getData().clear();
        if (selectedKind == null) {
            distributionChart.setTitle("分布は個別の処理を選択してください");
            return;
        }
        List<PipelineExecutionTimingSample> samples = tableSamples;
        if (samples.isEmpty()) {
            distributionChart.setTitle("サンプルなし");
            return;
        }
        int binCount = Math.min(20, Math.max(5, (int) Math.round(Math.sqrt(samples.size()) * 2)));
        List<PipelineExecutionTimingHistoryStore.HistogramBin> bins =
                PipelineExecutionTimingHistoryStore.computeHistogram(samples, binCount);
        XYChart.Series<String, Number> series = new XYChart.Series<>();
        series.setName(selectedKind.label());
        for (PipelineExecutionTimingHistoryStore.HistogramBin bin : bins) {
            series.getData().add(new XYChart.Data<>(bin.label(), bin.count()));
        }
        distributionChart.getData().add(series);
        distributionChart.setTitle(selectedKind.label() + " の分布（" + samples.size() + " 件）");
    }

    private void refreshStatsTable(PipelineExecutionTimingKind selectedKind, int limit) {
        List<StatsRow> rows = new ArrayList<>();
        if (selectedKind == null) {
            for (PipelineExecutionTimingKind kind : PipelineExecutionTimingKind.values()) {
                if (!isKindVisible(kind)) {
                    continue;
                }
                List<PipelineExecutionTimingSample> kindSamples =
                        shell.pipelineExecutionTimingHistory().recentSamplesForKind(kind, limit);
                rows.add(StatsRow.from(kind.label(), PipelineExecutionTimingHistoryStore.computeStats(kindSamples)));
            }
        } else {
            List<PipelineExecutionTimingSample> samples =
                    shell.pipelineExecutionTimingHistory().recentSamplesForKind(selectedKind, limit);
            rows.add(
                    StatsRow.from(
                            selectedKind.label(),
                            PipelineExecutionTimingHistoryStore.computeStats(samples)));
        }
        statsTable.setItems(FXCollections.observableArrayList(rows));
    }

    private record HistoryRow(
            String finishedAtText, String kindLabel, String durationText, String writerText) {}

    private record StatsRow(
            String kindLabel,
            String countText,
            String avgText,
            String medianText,
            String minText,
            String maxText,
            String stdDevText) {

        static StatsRow from(String kindLabel, Stats stats) {
            if (stats.count() <= 0L) {
                return new StatsRow(kindLabel, "—", "—", "—", "—", "—", "—");
            }
            return new StatsRow(
                    kindLabel,
                    Long.toString(stats.count()),
                    formatSec(stats.avgSec()),
                    formatSec(stats.medianSec()),
                    formatSec(stats.minSec()),
                    formatSec(stats.maxSec()),
                    formatSec(stats.stdDevSec()));
        }

        private static String formatSec(double sec) {
            return String.format(Locale.ROOT, "%.2f", sec);
        }
    }
}
