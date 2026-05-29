package jp.co.pm.ai.desktop;

import java.nio.file.Path;
import java.util.List;
import java.util.Locale;

import javafx.beans.property.ReadOnlyObjectWrapper;
import javafx.beans.property.ReadOnlyStringWrapper;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.chart.BarChart;
import javafx.scene.chart.XYChart;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.cell.PropertyValueFactory;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.dispatch.DispatchMlReadinessStore;

/** 学習速度分布・ML 有効化進捗の専用タブ。 */
public final class LearnedSpeedDataTabController {

    private MainShellController shell;

    @FXML private Label archivePathLabel;
    @FXML private Label speedStorePathLabel;
    @FXML private Label mlModeBadgeLabel;
    @FXML private Label mlArchiveProgressLabel;
    @FXML private ProgressBar mlArchiveProgressBar;
    @FXML private Label mlLadderLabel;
    @FXML private Label speedSummaryHeadlineLabel;
    @FXML private Label readinessUpdatedAtLabel;
    @FXML private TableView<SpeedKeyRow> speedKeyTable;
    @FXML private TableColumn<SpeedKeyRow, String> speedKeyColumn;
    @FXML private TableColumn<SpeedKeyRow, String> speedProcessColumn;
    @FXML private TableColumn<SpeedKeyRow, String> speedMachineColumn;
    @FXML private TableColumn<SpeedKeyRow, Number> speedSampleCountColumn;
    @FXML private TableColumn<SpeedKeyRow, String> speedP50Column;
    @FXML private TableColumn<SpeedKeyRow, String> speedAppliedColumn;
    @FXML private Label histogramDetailLabel;
    @FXML private BarChart<String, Number> speedHistogramChart;

    private final ObservableList<SpeedKeyRow> speedKeyRows = FXCollections.observableArrayList();

    @FXML
    private void initialize() {
        if (speedKeyTable != null) {
            speedKeyTable.setItems(speedKeyRows);
            wireSpeedKeyColumns();
            speedKeyTable
                    .getSelectionModel()
                    .selectedItemProperty()
                    .addListener((o, prev, cur) -> refreshHistogramForSelection(cur));
        }
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        refreshFromArchive();
    }

    void onMainShellTabSelected() {
        refreshFromArchive();
    }

    @FXML
    private void onRefreshAction() {
        refreshFromArchive();
    }

    void refreshFromArchive() {
        if (shell == null) {
            return;
        }
        Path archiveRoot = AppPaths.resolveDispatchLearningArchiveRoot(shell.snapshotUiEnv());
        Path speedStore =
                archiveRoot.resolve("speed-distributions").resolve("process_machine_speed.json");
        if (archivePathLabel != null) {
            archivePathLabel.setText("学習アーカイブ: " + archiveRoot);
        }
        if (speedStorePathLabel != null) {
            speedStorePathLabel.setText("速度分布 JSON: " + speedStore);
        }

        DispatchMlReadinessStore.ReadinessSnapshot readiness =
                DispatchMlReadinessStore.loadReadiness(archiveRoot);
        if (mlModeBadgeLabel != null) {
            mlModeBadgeLabel.setText("現在モード: " + readiness.mlModeActive());
        }
        if (mlArchiveProgressLabel != null) {
            mlArchiveProgressLabel.setText(
                    "アーカイブ job: " + readiness.archiveJobCount() + " / 20（ML-1）");
        }
        if (mlArchiveProgressBar != null) {
            mlArchiveProgressBar.setProgress(Math.min(1.0, readiness.archiveJobCount() / 20.0));
        }
        if (mlLadderLabel != null) {
            mlLadderLabel.setText(
                    "[MVP"
                            + (readiness.archiveJobCount() > 0 ? " ✓" : "")
                            + "] — [ML-0"
                            + (readiness.archiveJobCount() >= 10 ? " ✓" : "")
                            + "] — [ML-1"
                            + (readiness.archiveJobCount() >= 20 ? " ✓" : "")
                            + "] — [ML-2] — [ML-3]");
        }
        if (readinessUpdatedAtLabel != null) {
            String updated = readiness.updatedAt();
            readinessUpdatedAtLabel.setText(
                    updated != null && !updated.isBlank()
                            ? "ml_readiness 更新: " + updated
                            : "ml_readiness 未生成");
        }
        if (speedSummaryHeadlineLabel != null) {
            speedSummaryHeadlineLabel.setText(
                    "速度キー: "
                            + readiness.speedKeyCount()
                            + " 件（適用可能: "
                            + readiness.speedApplicableKeyCount()
                            + " 件）");
        }

        List<DispatchMlReadinessStore.SpeedKeyEntry> keys =
                DispatchMlReadinessStore.loadSpeedKeys(archiveRoot);
        SpeedKeyRow selected = speedKeyTable != null ? speedKeyTable.getSelectionModel().getSelectedItem() : null;
        String selectedKey = selected != null ? selected.key() : null;
        speedKeyRows.setAll(keys.stream().map(SpeedKeyRow::from).toList());
        if (speedKeyTable != null && !speedKeyRows.isEmpty()) {
            if (selectedKey != null) {
                speedKeyRows.stream()
                        .filter(r -> selectedKey.equals(r.key()))
                        .findFirst()
                        .ifPresentOrElse(
                                r -> speedKeyTable.getSelectionModel().select(r),
                                () -> speedKeyTable.getSelectionModel().selectFirst());
            } else {
                speedKeyTable.getSelectionModel().selectFirst();
            }
        } else {
            refreshHistogramForSelection(null);
        }
    }

    private void wireSpeedKeyColumns() {
        speedKeyColumn.setCellValueFactory(new PropertyValueFactory<>("key"));
        speedProcessColumn.setCellValueFactory(new PropertyValueFactory<>("process"));
        speedMachineColumn.setCellValueFactory(new PropertyValueFactory<>("machine"));
        speedSampleCountColumn.setCellValueFactory(new PropertyValueFactory<>("sampleCount"));
        speedP50Column.setCellValueFactory(new PropertyValueFactory<>("p50Text"));
        speedAppliedColumn.setCellValueFactory(new PropertyValueFactory<>("appliedText"));
    }

    private void refreshHistogramForSelection(SpeedKeyRow row) {
        if (histogramDetailLabel != null) {
            if (row == null) {
                histogramDetailLabel.setText("キーを選択するとヒストグラムを表示します。");
            } else {
                histogramDetailLabel.setText(
                        row.process()
                                + " / "
                                + row.machine()
                                + "  n="
                                + row.sampleCount()
                                + "  p50="
                                + row.p50Text()
                                + "  適用="
                                + row.appliedText());
            }
        }
        if (speedHistogramChart == null) {
            return;
        }
        speedHistogramChart.getData().clear();
        if (row == null || row.histogramCounts().isEmpty()) {
            return;
        }
        XYChart.Series<String, Number> series = new XYChart.Series<>();
        series.setName(row.key());
        List<Double> edges = row.histogramEdges();
        for (int i = 0; i < row.histogramCounts().size(); i++) {
            String label =
                    edges.size() > i + 1
                            ? String.format(Locale.ROOT, "%.0f-%.0f", edges.get(i), edges.get(i + 1))
                            : String.valueOf(i);
            series.getData().add(new XYChart.Data<>(label, row.histogramCounts().get(i)));
        }
        speedHistogramChart.getData().add(series);
    }

    public static final class SpeedKeyRow {
        private final ReadOnlyStringWrapper key = new ReadOnlyStringWrapper();
        private final ReadOnlyStringWrapper process = new ReadOnlyStringWrapper();
        private final ReadOnlyStringWrapper machine = new ReadOnlyStringWrapper();
        private final ReadOnlyObjectWrapper<Integer> sampleCount = new ReadOnlyObjectWrapper<>();
        private final ReadOnlyStringWrapper p50Text = new ReadOnlyStringWrapper();
        private final ReadOnlyStringWrapper appliedText = new ReadOnlyStringWrapper();
        private final List<Integer> histogramCounts;
        private final List<Double> histogramEdges;

        private SpeedKeyRow(
                String key,
                String process,
                String machine,
                int sampleCount,
                double p50,
                Double applied,
                List<Integer> histogramCounts,
                List<Double> histogramEdges) {
            this.key.set(key);
            this.process.set(process);
            this.machine.set(machine);
            this.sampleCount.set(sampleCount);
            this.p50Text.set(String.format(Locale.ROOT, "%.2f", p50));
            this.appliedText.set(applied != null ? String.format(Locale.ROOT, "%.2f", applied) : "—");
            this.histogramCounts = histogramCounts;
            this.histogramEdges = histogramEdges;
        }

        static SpeedKeyRow from(DispatchMlReadinessStore.SpeedKeyEntry entry) {
            return new SpeedKeyRow(
                    entry.key(),
                    entry.process(),
                    entry.machine(),
                    entry.n(),
                    entry.p50(),
                    entry.appliedSpeed(),
                    entry.histogramCounts(),
                    entry.histogramEdges());
        }

        public String getKey() {
            return key.get();
        }

        public String getProcess() {
            return process.get();
        }

        public String getMachine() {
            return machine.get();
        }

        public Integer getSampleCount() {
            return sampleCount.get();
        }

        public String getP50Text() {
            return p50Text.get();
        }

        public String getAppliedText() {
            return appliedText.get();
        }

        String key() {
            return getKey();
        }

        String process() {
            return getProcess();
        }

        String machine() {
            return getMachine();
        }

        int sampleCount() {
            return getSampleCount() != null ? getSampleCount() : 0;
        }

        String p50Text() {
            return getP50Text();
        }

        String appliedText() {
            return getAppliedText();
        }

        List<Integer> histogramCounts() {
            return histogramCounts;
        }

        List<Double> histogramEdges() {
            return histogramEdges;
        }
    }
}
