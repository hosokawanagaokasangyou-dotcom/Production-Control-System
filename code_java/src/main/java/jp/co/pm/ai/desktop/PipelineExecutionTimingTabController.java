package jp.co.pm.ai.desktop;

import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.EnumMap;
import java.util.LinkedHashMap;
import java.util.List;
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
import javafx.scene.control.TabPane;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;

import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * パイプライン実行時間のトレンド・分布（ヒストグラム）を表示するメインシェルタブ。
 */
public final class PipelineExecutionTimingTabController {

    private static final DateTimeFormatter FINISHED_AT =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss").withZone(ZoneId.systemDefault());

    /** トレンドで全種別を重ね表示するときのコンボ選択肢。 */
    private static final String KIND_FILTER_ALL = "（すべて・トレンド重ね）";

    private MainShellController shell;

    private final Runnable historyChangeListener = () -> Platform.runLater(this::refreshFromStore);

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
    private Label statsLabel;

    @FXML
    private TabPane chartTabPane;

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
        List<String> kindChoices = new ArrayList<>();
        kindChoices.add(KIND_FILTER_ALL);
        for (PipelineExecutionTimingKind kind : PipelineExecutionTimingKind.values()) {
            kindChoices.add(kind.label());
        }
        kindFilterCombo.setItems(FXCollections.observableArrayList(kindChoices));
        kindFilterCombo.getSelectionModel().selectFirst();
        kindFilterCombo
                .getSelectionModel()
                .selectedItemProperty()
                .addListener((o, a, b) -> refreshFromStore());

        sampleLimitSpinner.setValueFactory(
                new SpinnerValueFactory.IntegerSpinnerValueFactory(10, 500, 100, 10));
        sampleLimitSpinner.valueProperty().addListener((o, a, b) -> refreshFromStore());

        historyTimeColumn.setCellValueFactory(
                cd -> new ReadOnlyStringWrapper(cd.getValue().finishedAtText()));
        historyKindColumn.setCellValueFactory(
                cd -> new ReadOnlyStringWrapper(cd.getValue().kindLabel()));
        historyDurationColumn.setCellValueFactory(
                cd -> new ReadOnlyStringWrapper(cd.getValue().durationText()));
        historyWriterColumn.setCellValueFactory(
                cd -> new ReadOnlyStringWrapper(cd.getValue().writerText()));

        for (PipelineExecutionTimingKind kind : PipelineExecutionTimingKind.values()) {
            XYChart.Series<Number, Number> series = new XYChart.Series<>();
            series.setName(kind.label());
            trendSeriesByKind.put(kind, series);
            trendChart.getData().add(series);
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
        if (hintLabel != null) {
            PipelineExecutionTimingHistoryStore store = shell.pipelineExecutionTimingHistory();
            String lockHint =
                    store.isPersistLocked()
                            ? store.readPersistLockInfo()
                                    .map(
                                            info ->
                                                    " 他端末が保存中: "
                                                            + info.displayHost()
                                                            + " / "
                                                            + info.displayIp())
                                    .orElse(" 他端末が保存中")
                            : "";
            hintLabel.setText(
                    "段階1～3・サマリ Excel・納期管理ビューの実行時間を記録します。"
                            + " 履歴ファイルはサマリ Excel（"
                            + AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK
                            + "）と同一フォルダに保存されます（保存前に .save.lock で排他）。"
                            + lockHint
                            + " ファイル: "
                            + store.storagePath());
        }
        PipelineExecutionTimingHistoryStore store = shell.pipelineExecutionTimingHistory();
        int limit = sampleLimitSpinner.getValue() != null ? sampleLimitSpinner.getValue() : 100;
        PipelineExecutionTimingKind selectedKind = selectedKindOrNull();

        List<PipelineExecutionTimingSample> tableSamples =
                selectedKind == null
                        ? store.recentSamples(null, limit)
                        : store.recentSamplesForKind(selectedKind, limit);
        refreshHistoryTable(tableSamples);
        refreshTrendChart(store, selectedKind, limit);
        refreshDistributionChart(selectedKind, limit, tableSamples);
        refreshStatsLabel(tableSamples);
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

    private void refreshStatsLabel(List<PipelineExecutionTimingSample> samples) {
        PipelineExecutionTimingKind selectedKind = selectedKindOrNull();
        if (selectedKind == null) {
            Map<PipelineExecutionTimingKind, PipelineExecutionTimingHistoryStore.Stats> byKind =
                    new LinkedHashMap<>();
            int limit = sampleLimitSpinner.getValue() != null ? sampleLimitSpinner.getValue() : 100;
            for (PipelineExecutionTimingKind kind : PipelineExecutionTimingKind.values()) {
                List<PipelineExecutionTimingSample> kindSamples =
                        shell.pipelineExecutionTimingHistory().recentSamplesForKind(kind, limit);
                byKind.put(kind, PipelineExecutionTimingHistoryStore.computeStats(kindSamples));
            }
            StringBuilder sb = new StringBuilder();
            for (Map.Entry<PipelineExecutionTimingKind, PipelineExecutionTimingHistoryStore.Stats> e :
                    byKind.entrySet()) {
                if (sb.length() > 0) {
                    sb.append('\n');
                }
                sb.append(e.getKey().label())
                        .append(": ")
                        .append(e.getValue().summaryJa());
            }
            statsLabel.setText(sb.toString());
            return;
        }
        statsLabel.setText(PipelineExecutionTimingHistoryStore.computeStats(samples).summaryJa());
    }

    private record HistoryRow(
            String finishedAtText, String kindLabel, String durationText, String writerText) {}
}
