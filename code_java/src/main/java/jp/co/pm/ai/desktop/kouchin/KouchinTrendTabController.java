package jp.co.pm.ai.desktop.kouchin;

import java.awt.Desktop;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import javafx.application.Platform;
import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.scene.chart.LineChart;
import javafx.scene.chart.NumberAxis;
import javafx.scene.chart.XYChart;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;

import jp.co.pm.ai.desktop.MainShellController;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.io.DesktopFileOpener;
import jp.co.pm.ai.kouchin.trend.TrendResult;
import jp.co.pm.ai.kouchin.trend.TrendModel;
import jp.co.pm.ai.kouchin.trend.TrendRequest;
import jp.co.pm.ai.kouchin.trend.TrendService;
import jp.co.pm.ai.kouchin.trend.TrendShiftRow;
import jp.co.pm.ai.kouchin.verify.FileDiscovery;
import jp.co.pm.ai.kouchin.verify.KouchinPaths;
import jp.co.pm.ai.kouchin.verify.YearMonthKey;

/**
 * 後加工工賃の月次トレンド。加工トレンドタブの日次グラフとは別。集計データを混ぜない。
 */
public class KouchinTrendTabController {

    @FXML private Spinner<Integer> monthsSpinner;
    @FXML private ComboBox<String> factoryCombo;
    @FXML private Button runButton;
    @FXML private Button openExcelButton;
    @FXML private Label statusLabel;
    @FXML private Label summaryLabel;
    @FXML private LineChart<Number, Number> volumeChart;
    @FXML private LineChart<Number, Number> feeChart;
    @FXML private NumberAxis volumeXAxis;
    @FXML private NumberAxis feeXAxis;

    private MainShellController shell;
    private KouchinHostTabController host;
    private TrendResult last;
    private boolean selected;

    @FXML
    private void initialize() {
        if (monthsSpinner != null) {
            monthsSpinner.setValueFactory(new SpinnerValueFactory.IntegerSpinnerValueFactory(2, 24, 6));
        }
        if (factoryCombo != null) {
            factoryCombo.getItems().setAll("両方", "国分", "湖南");
            factoryCombo.getSelectionModel().selectFirst();
        }
        styleChart(volumeChart);
        styleChart(feeChart);
        if (statusLabel != null) {
            statusLabel.setText("まだ月次トレンドを作成していません");
        }
        refreshRunEnabled();
    }

    public void bindShell(MainShellController shell, KouchinHostTabController host) {
        this.shell = shell;
        this.host = host;
        refreshRunEnabled();
    }

    public void onMainShellTabSelected() {
        selected = true;
        refreshRunEnabled();
    }

    public void onMainShellTabDeselected() {
        selected = false;
    }

    public void refreshRunEnabled() {
        boolean unapplied = host != null && host.hasUnappliedSourceEdits();
        boolean busy = shell != null && (shell.isKouchinRunBusy() || shell.isPlanningPipelineStageRunning());
        if (runButton != null) {
            runButton.setDisable(unapplied || busy);
        }
    }

    @FXML
    private void onRunTrend() {
        if (shell == null) {
            return;
        }
        if (host != null && host.hasUnappliedSourceEdits()) {
            statusLabel.setText("参照先が未適用のため実行できません");
            return;
        }
        if (!shell.tryBeginKouchinRun("後加工工賃 月次トレンド作成中…")) {
            return;
        }
        refreshRunEnabled();
        Map<String, String> ui = shell.snapshotUiEnv();
        KouchinPaths paths = KouchinPaths.fromEnv(ui);
        int months = monthsSpinner == null || monthsSpinner.getValue() == null ? 6 : monthsSpinner.getValue();
        String fac = factoryCode();
        Task<TrendResult> task = new Task<>() {
            @Override
            protected TrendResult call() {
                FileDiscovery.invalidateListingCache();
                return TrendService.run(new TrendRequest(months, fac, paths, ui));
            }
        };
        task.setOnSucceeded(e -> {
            shell.endKouchinRun();
            last = task.getValue();
            show(last);
            refreshRunEnabled();
        });
        task.setOnFailed(e -> {
            shell.endKouchinRun();
            Throwable err = task.getException();
            if (statusLabel != null) {
                statusLabel.setText("失敗: " + (err == null ? "" : err.getMessage()));
            }
            refreshRunEnabled();
        });
        Thread t = new Thread(task, "kouchin-trend");
        t.setDaemon(true);
        t.start();
    }

    @FXML
    private void onOpenExcel() {
        Path p = last == null ? null : last.preferredExcel();
        if (p == null && last != null && last.xlsx() != null && !last.xlsx().succeeded().isEmpty()) {
            FactorySite site = shell == null ? FactorySite.KOKUBU : shell.currentFactorySite();
            Path prefer = site == FactorySite.KONAN
                    ? last.xlsx().succeeded().get(Math.min(1, last.xlsx().succeeded().size() - 1))
                    : last.xlsx().succeeded().get(0);
            p = prefer;
        }
        if (p == null) {
            return;
        }
        try {
            DesktopFileOpener.openFile(p);
            statusLabel.setText("開いた: " + p.toAbsolutePath());
        } catch (Exception e) {
            try {
                Desktop.getDesktop().open(p.toFile());
            } catch (Exception ex) {
                statusLabel.setText("開けません: " + ex.getMessage());
            }
        }
    }

    private void show(TrendResult result) {
        if (result == null) {
            return;
        }
        TrendModel model = result.model();
        StringBuilder sum = new StringBuilder();
        if (model != null) {
            sum.append("注目工程: ").append(String.join(", ", model.focusProcesses()));
            sum.append("  振替疑い: ");
            int n = 0;
            for (TrendShiftRow s : model.suspects()) {
                if (n++ >= 3) {
                    break;
                }
                sum.append(s.process()).append("(").append(s.badge()).append(") ");
            }
            fillChart(volumeChart, model, "qty");
            fillChart(feeChart, model, "wage");
        }
        if (summaryLabel != null) {
            summaryLabel.setText(sum.toString());
        }
        StringBuilder st = new StringBuilder("作成完了。");
        if (result.xlsx() != null) {
            st.append(" 成功 ").append(result.xlsx().succeeded().size());
            if (result.xlsx().anyFailed()) {
                st.append(" 失敗 ").append(String.join(" / ", result.xlsx().failures()));
            }
        }
        if (statusLabel != null) {
            statusLabel.setText(st.toString());
        }
        if (shell != null) {
            shell.appendLog("[kouchin-trend] " + st);
        }
    }

    private void fillChart(LineChart<Number, Number> chart, TrendModel model, String metric) {
        if (chart == null || model == null) {
            return;
        }
        chart.getData().clear();
        var matrix = "qty".equals(metric) ? model.qty() : model.wage();
        List<YearMonthKey> months = model.months();
        List<String> focus = model.focusProcesses();
        if (focus.isEmpty()) {
            return;
        }
        String proc = focus.get(0);
        XYChart.Series<Number, Number> k = new XYChart.Series<>();
        k.setName("国分");
        XYChart.Series<Number, Number> n = new XYChart.Series<>();
        n.setName("湖南");
        for (int i = 0; i < months.size(); i++) {
            YearMonthKey ym = months.get(i);
            Double kv = matrix.values().getOrDefault("国分", Map.of()).getOrDefault(proc, Map.of()).get(ym);
            Double nv = matrix.values().getOrDefault("湖南", Map.of()).getOrDefault(proc, Map.of()).get(ym);
            if (kv != null) {
                k.getData().add(new XYChart.Data<>(i + 1, kv));
            }
            if (nv != null) {
                n.getData().add(new XYChart.Data<>(i + 1, nv));
            }
        }
        chart.getData().addAll(k, n);
        Platform.runLater(() -> {
            if (k.getNode() != null) {
                k.getNode().setStyle("-fx-stroke: #C0504D;");
            }
            if (n.getNode() != null) {
                n.getNode().setStyle("-fx-stroke: #9BBB59;");
            }
        });
    }

    private static void styleChart(LineChart<Number, Number> chart) {
        if (chart == null) {
            return;
        }
        chart.setCreateSymbols(true);
        chart.setAnimated(false);
    }

    private String factoryCode() {
        String v = factoryCombo == null ? "両方" : factoryCombo.getValue();
        if ("国分".equals(v)) {
            return TrendRequest.FACTORY_KOKUBU;
        }
        if ("湖南".equals(v)) {
            return TrendRequest.FACTORY_KONAN;
        }
        return TrendRequest.FACTORY_BOTH;
    }
}
