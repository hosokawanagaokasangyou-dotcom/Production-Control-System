package jp.co.pm.ai.desktop;

import java.lang.management.ManagementFactory;
import java.lang.management.MemoryMXBean;
import java.lang.management.MemoryUsage;
import java.util.Locale;
import java.util.concurrent.atomic.AtomicBoolean;

import javafx.animation.KeyFrame;
import javafx.animation.Timeline;
import javafx.fxml.FXML;
import javafx.scene.chart.LineChart;
import javafx.scene.chart.NumberAxis;
import javafx.scene.chart.XYChart;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.Label;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.layout.VBox;
import javafx.util.Duration;

import jp.co.pm.ai.desktop.config.DesktopSessionState;

/**
 * JVM ヒープの参照・次回起動時の希望値、およびヒープ使用量の定期監視と推移グラフ。
 */
public final class MemorySettingsTabController {

    private static final int MAX_CHART_POINTS = 400;

    private MainShellController shell;

    @FXML
    private Label heapSectionTitleLabel;

    @FXML
    private Label heapExplainLabel;

    @FXML
    private Label desiredHeapLabel;

    @FXML
    private Spinner<Integer> nextLaunchHeapMiBSpinner;

    @FXML
    private Button syncHeapToCurrentButton;

    @FXML
    private Label heapArgHintLabel;

    @FXML
    private Label monitorSectionTitleLabel;

    @FXML
    private CheckBox monitorEnabledCheck;

    @FXML
    private Label intervalCaptionLabel;

    @FXML
    private Spinner<Integer> intervalSpinner;

    @FXML
    private Label monitorHintLabel;

    @FXML
    private Label chartCaptionLabel;

    @FXML
    private VBox chartContainer;

    @FXML
    private LineChart<Number, Number> heapChart;

    @FXML
    private Label runtimeCaptionLabel;

    @FXML
    private Label runtimeSlashLabel;

    @FXML
    private Label runtimeOpenParenLabel;

    @FXML
    private Label runtimeHeapUsedLabel;

    @FXML
    private Label runtimeHeapMaxLabel;

    @FXML
    private Label runtimeHeapCommittedLabel;

    @FXML
    private Label runtimeCloseParenLabel;

    @FXML
    private Button refreshRuntimeHeapButton;

    private final XYChart.Series<Number, Number> heapUsedSeries = new XYChart.Series<>();

    private Timeline monitorTimeline;

    private long monitorStartedEpochMs;

    private final AtomicBoolean suppressPersist = new AtomicBoolean(false);

    private boolean listenersInstalled;

    @FXML
    private void initialize() {
        heapSectionTitleLabel.setText("ヒープサイズ（次回 JVM 起動時）");
        heapExplainLabel.setText(
                "実行中の JVM はヒープ上限（-Xmx）を変更できません。"
                        + " ここで指定した値はセッションに保存され、次回起動時の JVM 引数の参考として使えます。");
        desiredHeapLabel.setText("希望ヒープ上限（MiB）");
        monitorSectionTitleLabel.setText("メモリ監視");
        monitorEnabledCheck.setText("ヒープ使用量を監視する");
        intervalCaptionLabel.setText("間隔（秒）");
        chartCaptionLabel.setText("ヒープ使用量の推移");
        runtimeCaptionLabel.setText("現在のヒープ（実行中 JVM）");
        runtimeSlashLabel.setText("/");
        runtimeOpenParenLabel.setText("（コミット ");
        runtimeCloseParenLabel.setText("）");
        monitorHintLabel.setText(
                "監視をオンにすると、指定間隔でヒープ使用量を記録し、下のグラフに表示します。"
                        + " オフにするとタイマーを停止し、グラフを隠します。");
        syncHeapToCurrentButton.setText("現在の上限に合わせる");
        refreshRuntimeHeapButton.setText("今すぐ更新");

        heapUsedSeries.setName("ヒープ使用量");
        heapChart.getData().add(heapUsedSeries);
        heapChart.setCreateSymbols(false);
        NumberAxis xAxis = (NumberAxis) heapChart.getXAxis();
        xAxis.setLabel("経過（秒）");
        NumberAxis yAxis = (NumberAxis) heapChart.getYAxis();
        yAxis.setLabel("使用量（MiB）");

        int curMaxMiB = readHeapMaxMiBOrFallback();
        nextLaunchHeapMiBSpinner.setValueFactory(
                new SpinnerValueFactory.IntegerSpinnerValueFactory(128, 65_536, curMaxMiB, 256));

        intervalSpinner.setValueFactory(new SpinnerValueFactory.IntegerSpinnerValueFactory(1, 3600, 5));

        refreshRuntimeHeapLabels();
        updateHeapArgHint();

        nextLaunchHeapMiBSpinner
                .valueProperty()
                .addListener((obs, o, n) -> {
                    updateHeapArgHint();
                    persistIfReady();
                });

        chartContainer.setManaged(false);
        chartContainer.setVisible(false);
        heapChart.setManaged(false);
        heapChart.setVisible(false);
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        if (!listenersInstalled) {
            listenersInstalled = true;
            monitorEnabledCheck
                    .selectedProperty()
                    .addListener(
                            (obs, o, n) -> {
                                restartMonitorFromUiState();
                                persistIfReady();
                            });
            intervalSpinner
                    .valueProperty()
                    .addListener(
                            (obs, o, n) -> {
                                restartMonitorFromUiState();
                                persistIfReady();
                            });
        }
    }

    /** セッション適用（{@link MainShellController#applyDesktopSession} から）。 */
    void applyMemorySettingsSession(DesktopSessionState s) {
        if (s == null) {
            return;
        }
        suppressPersist.set(true);
        try {
            monitorEnabledCheck.setSelected(s.memoryMonitorEnabled());
            int iv = (int) clamp(s.memoryMonitorIntervalSec(), 1, 3600);
            intervalSpinner.getValueFactory().setValue(iv);

            long saved = s.nextLaunchHeapMaxMiB();
            int cur = readHeapMaxMiBOrFallback();
            int target = saved > 0 ? (int) clamp(saved, 128, 65_536) : cur;
            nextLaunchHeapMiBSpinner.getValueFactory().setValue(target);

            updateHeapArgHint();
        } finally {
            suppressPersist.set(false);
        }
        restartMonitorFromUiState();
    }

    boolean snapshotMemoryMonitorEnabled() {
        return monitorEnabledCheck.isSelected();
    }

    long snapshotMemoryMonitorIntervalSec() {
        return clamp(intervalSpinner.getValue(), 1, 3600);
    }

    long snapshotNextLaunchHeapMaxMiB() {
        int v = nextLaunchHeapMiBSpinner.getValue();
        return v <= 0 ? 0L : v;
    }

    /** メインウィンドウ終了時にタイマーを停止する。 */
    void shutdown() {
        stopMonitorTimeline();
    }

    @FXML
    private void onSyncHeapToCurrentAction() {
        int cur = readHeapMaxMiBOrFallback();
        nextLaunchHeapMiBSpinner.getValueFactory().setValue(cur);
        updateHeapArgHint();
        persistIfReady();
    }

    @FXML
    private void onRefreshRuntimeHeapAction() {
        refreshRuntimeHeapLabels();
    }

    private void persistIfReady() {
        if (suppressPersist.get() || shell == null) {
            return;
        }
        shell.persistDesktopSessionNow();
    }

    private void restartMonitorFromUiState() {
        stopMonitorTimeline();
        boolean on = monitorEnabledCheck.isSelected();
        chartContainer.setManaged(on);
        chartContainer.setVisible(on);
        heapChart.setManaged(on);
        heapChart.setVisible(on);
        if (!on) {
            heapUsedSeries.getData().clear();
            return;
        }
        heapUsedSeries.getData().clear();
        monitorStartedEpochMs = System.currentTimeMillis();
        int sec = (int) clamp(intervalSpinner.getValue(), 1, 3600);
        KeyFrame kf =
                new KeyFrame(
                        Duration.seconds(sec),
                        event -> {
                            appendHeapSample();
                            refreshRuntimeHeapLabels();
                        });
        monitorTimeline = new Timeline(kf);
        monitorTimeline.setCycleCount(Timeline.INDEFINITE);
        monitorTimeline.play();
        appendHeapSample();
    }

    private void appendHeapSample() {
        MemoryMXBean bean = ManagementFactory.getMemoryMXBean();
        MemoryUsage heap = bean.getHeapMemoryUsage();
        double tSec = (System.currentTimeMillis() - monitorStartedEpochMs) / 1000.0;
        long usedMiB = heap.getUsed() / (1024L * 1024L);
        heapUsedSeries.getData().add(new XYChart.Data<>(tSec, usedMiB));
        while (heapUsedSeries.getData().size() > MAX_CHART_POINTS) {
            heapUsedSeries.getData().removeFirst();
        }
    }

    private void stopMonitorTimeline() {
        if (monitorTimeline != null) {
            monitorTimeline.stop();
            monitorTimeline = null;
        }
    }

    private void refreshRuntimeHeapLabels() {
        MemoryMXBean bean = ManagementFactory.getMemoryMXBean();
        MemoryUsage heap = bean.getHeapMemoryUsage();
        long used = heap.getUsed();
        long max = heap.getMax();
        long committed = heap.getCommitted();
        runtimeHeapUsedLabel.setText(formatMiB(used));
        runtimeHeapMaxLabel.setText(formatMiBMax(max));
        runtimeHeapCommittedLabel.setText(formatMiB(committed));
    }

    private void updateHeapArgHint() {
        int mib = nextLaunchHeapMiBSpinner.getValue();
        heapArgHintLabel.setText(
                String.format(Locale.ROOT, "例: java -Xmx%dm ... （起動スクリプトや IDE の VM オプションに設定）", mib));
    }

    private static String formatMiB(long bytes) {
        if (bytes < 0) {
            return "n/a";
        }
        return String.format(Locale.ROOT, "%d MiB", bytes / (1024L * 1024L));
    }

    private static String formatMiBMax(long maxBytes) {
        if (maxBytes < 0) {
            return "n/a";
        }
        return formatMiB(maxBytes);
    }

    private static int readHeapMaxMiBOrFallback() {
        long max = ManagementFactory.getMemoryMXBean().getHeapMemoryUsage().getMax();
        if (max <= 0) {
            return 2048;
        }
        return (int) Math.min(65_536L, Math.max(128L, max / (1024L * 1024L)));
    }

    private static long clamp(long v, long lo, long hi) {
        return Math.max(lo, Math.min(hi, v));
    }
}
