package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.lang.management.ManagementFactory;
import java.lang.management.MemoryMXBean;
import java.lang.management.MemoryUsage;
import java.nio.charset.StandardCharsets;
import java.text.MessageFormat;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Properties;
import java.util.concurrent.atomic.AtomicBoolean;

import javafx.animation.KeyFrame;
import javafx.animation.Timeline;
import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.scene.chart.LineChart;
import javafx.scene.chart.NumberAxis;
import javafx.scene.chart.XYChart;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.Label;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.TextArea;
import javafx.scene.layout.VBox;
import javafx.util.Duration;

import jp.co.pm.ai.desktop.config.DesktopSessionState;
import jp.co.pm.ai.desktop.config.JvmMemoryLogStore;
import jp.co.pm.ai.desktop.config.PomJvmHeapPropertiesSync;
import jp.co.pm.ai.desktop.runtime.MemoryJvmRingLog;

/**
 * Memory settings tab. UI strings load from classpath {@code memory_settings_strings.properties}
 * (ASCII + \\u escapes) so Japanese is not stored in this source file.
 */
public final class MemorySettingsTabController {

    private static final String STRINGS_RESOURCE = "/jp/co/pm/ai/desktop/memory_settings_strings.properties";

    private static final Properties UI_STRINGS = loadStrings();

    private static Properties loadStrings() {
        Properties p = new Properties();
        try (InputStream in = MemorySettingsTabController.class.getResourceAsStream(STRINGS_RESOURCE)) {
            Objects.requireNonNull(in, STRINGS_RESOURCE);
            try (InputStreamReader r = new InputStreamReader(in, StandardCharsets.UTF_8)) {
                p.load(r);
            }
        } catch (IOException e) {
            throw new ExceptionInInitializerError(e);
        }
        return p;
    }

    private static String txt(String key) {
        String v = UI_STRINGS.getProperty(key);
        return v != null ? v : key;
    }

    private static String txtFmt(String key, Object... args) {
        return MessageFormat.format(txt(key), args);
    }

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

    @FXML
    private Label jvmLogSectionTitleLabel;

    @FXML
    private Label jvmLogMaxLinesLabel;

    @FXML
    private Spinner<Integer> memoryJvmLogMaxLinesSpinner;

    @FXML
    private Label jvmLogHintLabel;

    @FXML
    private TextArea jvmLogTextArea;

    private final XYChart.Series<Number, Number> heapUsedSeries = new XYChart.Series<>();

    private Timeline monitorTimeline;

    private long monitorStartedEpochMs;

    private final AtomicBoolean suppressPersist = new AtomicBoolean(false);

    private boolean listenersInstalled;

    @FXML
    private void initialize() {
        heapSectionTitleLabel.setText(txt("heap.section.title"));
        heapExplainLabel.setText(txt("heap.explain"));
        desiredHeapLabel.setText(txt("heap.desired.label"));
        monitorSectionTitleLabel.setText(txt("monitor.section.title"));
        monitorEnabledCheck.setText(txt("monitor.enabled"));
        intervalCaptionLabel.setText(txt("interval.caption"));
        chartCaptionLabel.setText(txt("chart.caption"));
        runtimeCaptionLabel.setText(txt("runtime.caption"));
        runtimeSlashLabel.setText("/");
        runtimeOpenParenLabel.setText(txt("runtime.open_paren"));
        runtimeCloseParenLabel.setText(txt("runtime.close_paren"));
        monitorHintLabel.setText(txt("monitor.hint"));
        syncHeapToCurrentButton.setText(txt("sync.button"));
        refreshRuntimeHeapButton.setText(txt("refresh.button"));

        heapUsedSeries.setName(txt("series.heap_used"));
        heapChart.getData().add(heapUsedSeries);
        heapChart.setCreateSymbols(false);
        NumberAxis xAxis = (NumberAxis) heapChart.getXAxis();
        xAxis.setLabel(txt("axis.elapsed_sec"));
        NumberAxis yAxis = (NumberAxis) heapChart.getYAxis();
        yAxis.setLabel(txt("axis.usage_mib"));

        int curMaxMiB = readHeapMaxMiBOrFallback();
        nextLaunchHeapMiBSpinner.setValueFactory(
                new SpinnerValueFactory.IntegerSpinnerValueFactory(
                        PomJvmHeapPropertiesSync.MIN_HEAP_MIB, 65_536, curMaxMiB, 256));

        intervalSpinner.setValueFactory(new SpinnerValueFactory.IntegerSpinnerValueFactory(1, 3600, 5));

        jvmLogSectionTitleLabel.setText(txt("jvm.log.section"));
        jvmLogMaxLinesLabel.setText(txt("jvm.log.max_lines"));
        jvmLogHintLabel.setText(txt("jvm.log.hint"));
        memoryJvmLogMaxLinesSpinner.setValueFactory(
                new SpinnerValueFactory.IntegerSpinnerValueFactory(
                        MemoryJvmRingLog.ABS_MIN,
                        MemoryJvmRingLog.ABS_MAX,
                        MemoryJvmRingLog.DEFAULT_MAX_LINES,
                        500));
        memoryJvmLogMaxLinesSpinner
                .valueProperty()
                .addListener(
                        (obs, o, n) -> {
                            if (n != null) {
                                MemoryJvmRingLog.setMaxLines(n);
                                persistIfReady();
                            }
                        });

        refreshRuntimeHeapLabels();
        updateHeapArgHint();

        nextLaunchHeapMiBSpinner
                .valueProperty()
                .addListener(
                        (obs, o, n) -> {
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
        MemoryJvmRingLog.setUiRefreshListener(
                () -> Platform.runLater(this::flushJvmLogTextArea));
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

    /** Applies persisted session (invoked from {@link MainShellController#applyDesktopSession}). */
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
            int target =
                    saved > 0
                            ? (int) clamp(saved, PomJvmHeapPropertiesSync.MIN_HEAP_MIB, 65_536)
                            : cur;
            nextLaunchHeapMiBSpinner.getValueFactory().setValue(target);

            memoryJvmLogMaxLinesSpinner.getValueFactory().setValue(MemoryJvmRingLog.getMaxLines());

            updateHeapArgHint();
        } finally {
            suppressPersist.set(false);
        }
        restartMonitorFromUiState();
        flushJvmLogTextArea();
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

    /** Stops timers and clears log listener on main window close. */
    void shutdown() {
        stopMonitorTimeline();
        MemoryJvmRingLog.setUiRefreshListener(null);
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
        PomJvmHeapPropertiesSync.writeJvmHeapFromDesiredMiB(
                shell.snapshotUiEnv(), nextLaunchHeapMiBSpinner.getValue());
        JvmMemoryLogStore.persistSnapshot(
                memoryJvmLogMaxLinesSpinner.getValue(), MemoryJvmRingLog.snapshotLines());
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
        heapArgHintLabel.setText(txtFmt("heap.arg.hint", mib));
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
        return (int)
                Math.min(
                        65_536L,
                        Math.max((long) PomJvmHeapPropertiesSync.MIN_HEAP_MIB, max / (1024L * 1024L)));
    }

    private static long clamp(long v, long lo, long hi) {
        return Math.max(lo, Math.min(hi, v));
    }

    private void flushJvmLogTextArea() {
        if (jvmLogTextArea == null) {
            return;
        }
        List<String> lines = MemoryJvmRingLog.snapshotLines();
        jvmLogTextArea.setText(String.join("\n", lines));
        int len = jvmLogTextArea.getLength();
        jvmLogTextArea.positionCaret(len);
        jvmLogTextArea.setScrollTop(Double.MAX_VALUE);
    }
}
