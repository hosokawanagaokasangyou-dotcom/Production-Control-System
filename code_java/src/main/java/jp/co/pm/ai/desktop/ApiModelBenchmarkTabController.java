package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicReference;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextArea;
import javafx.scene.control.cell.PropertyValueFactory;

import jp.co.pm.ai.desktop.benchmark.GeminiGenerateContentRestClient;
import jp.co.pm.ai.desktop.benchmark.GeminiGenerateContentRestClient.CallResult;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.GeminiDispatchModelTryOrderDefaults;
import jp.co.pm.ai.desktop.crypto.GeminiCredentialsV2Crypto;

/**
 * Gemini {@code generateContent} の往復レイテンシを計測するメインシェルタブ。
 */
public final class ApiModelBenchmarkTabController {

    private static final List<String> DEFAULT_MODELS =
            List.copyOf(GeminiDispatchModelTryOrderDefaults.PLANNING_CORE_FALLBACK_TRY_ORDER);

    private MainShellController shell;

    @FXML
    private Label credentialsPathLabel;

    @FXML
    private Label modelPriorityLabel;

    @FXML
    private ProgressBar benchProgressBar;

    @FXML
    private Label benchRunningLabel;

    @FXML
    private Button cancelBenchButton;

    @FXML
    private CheckBox completionNotifyCheck;

    @FXML
    private ComboBox<String> modelCombo;

    @FXML
    private Spinner<Integer> runsSpinner;

    @FXML
    private Spinner<Integer> maxTokensSpinner;

    @FXML
    private Spinner<Integer> timeoutSecondsSpinner;

    @FXML
    private CheckBox warmupCheck;

    @FXML
    private CheckBox stopOnTerminalHttpCheck;

    @FXML
    private Button runButton;

    @FXML
    private Button runAllModelsButton;

    @FXML
    private Button clearButton;

    @FXML
    private TextArea promptArea;

    @FXML
    private TextArea logArea;

    @FXML
    private TableView<BenchmarkTableRow> rankingTable;

    @FXML
    private Label summaryLabel;

    private final AtomicReference<Thread> worker = new AtomicReference<>();

    /** 中断ボタンで立てる。ワーカースレッドがポーリングする。 */
    private final AtomicBoolean benchmarkCancelRequested = new AtomicBoolean(false);

    private record ModelBenchOutcome(
            String model,
            int http2xxCount,
            int measuredAttempts,
            int plannedRuns,
            double avgMs,
            double minMs,
            double maxMs,
            String supplement,
            boolean userInterrupted) {

        String toSummaryLine() {
            StringBuilder sb = new StringBuilder();
            if (userInterrupted) {
                sb.append(model).append(": ユーザー中断（実施 ").append(measuredAttempts).append("/").append(plannedRuns).append("）");
            } else if (http2xxCount <= 0) {
                sb.append(model)
                        .append(": HTTP 2xx なし（実施 ")
                        .append(measuredAttempts)
                        .append("/")
                        .append(plannedRuns)
                        .append("）");
            } else {
                sb.append(
                        String.format(
                                Locale.ROOT,
                                "%s: 平均 %.2f ms / 最小 %.2f / 最大 %.2f（n=%d）",
                                model, avgMs, minMs, maxMs, http2xxCount));
            }
            if (supplement != null && !supplement.isBlank()) {
                sb.append(" — ").append(supplement.strip());
            }
            return sb.toString();
        }
    }

    /** ランキング表の1行（JavaFX {@link PropertyValueFactory} 用ゲッター）。 */
    public static final class BenchmarkTableRow {
        private final int rank;
        private final String modelName;
        private final String avgMillis;
        private final String minMillis;
        private final String maxMillis;
        private final int successCount;
        private final String attempts;
        private final String notes;

        public BenchmarkTableRow(
                int rank,
                String modelName,
                String avgMillis,
                String minMillis,
                String maxMillis,
                int successCount,
                String attempts,
                String notes) {
            this.rank = rank;
            this.modelName = modelName;
            this.avgMillis = avgMillis;
            this.minMillis = minMillis;
            this.maxMillis = maxMillis;
            this.successCount = successCount;
            this.attempts = attempts;
            this.notes = notes;
        }

        public int getRank() {
            return rank;
        }

        public String getModelName() {
            return modelName;
        }

        public String getAvgMillis() {
            return avgMillis;
        }

        public String getMinMillis() {
            return minMillis;
        }

        public String getMaxMillis() {
            return maxMillis;
        }

        public int getSuccessCount() {
            return successCount;
        }

        public String getAttempts() {
            return attempts;
        }

        public String getNotes() {
            return notes;
        }

        static BenchmarkTableRow ranked(int rank, ModelBenchOutcome o) {
            String avg =
                    o.http2xxCount() > 0 && !Double.isNaN(o.avgMs())
                            ? String.format(Locale.ROOT, "%.2f", o.avgMs())
                            : "—";
            String min =
                    o.http2xxCount() > 0 && !Double.isNaN(o.minMs())
                            ? String.format(Locale.ROOT, "%.2f", o.minMs())
                            : "—";
            String max =
                    o.http2xxCount() > 0 && !Double.isNaN(o.maxMs())
                            ? String.format(Locale.ROOT, "%.2f", o.maxMs())
                            : "—";
            String att = o.measuredAttempts() + " / " + o.plannedRuns();
            String notes = buildNotes(o);
            return new BenchmarkTableRow(rank, o.model(), avg, min, max, o.http2xxCount(), att, notes);
        }

        private static String buildNotes(ModelBenchOutcome o) {
            StringBuilder n = new StringBuilder();
            if (o.userInterrupted()) {
                n.append("中断");
            }
            if (o.supplement() != null && !o.supplement().isBlank()) {
                if (n.length() > 0) {
                    n.append(" · ");
                }
                n.append(o.supplement().strip());
            }
            return n.length() > 0 ? n.toString() : "—";
        }
    }

    @FXML
    private void initialize() {
        if (modelCombo != null) {
            modelCombo.setItems(FXCollections.observableArrayList(DEFAULT_MODELS));
            modelCombo.getSelectionModel().selectFirst();
        }
        if (runsSpinner != null) {
            runsSpinner.setValueFactory(new SpinnerValueFactory.IntegerSpinnerValueFactory(1, 100, 5, 1));
        }
        if (maxTokensSpinner != null) {
            maxTokensSpinner.setValueFactory(new SpinnerValueFactory.IntegerSpinnerValueFactory(8, 2048, 64, 8));
        }
        if (timeoutSecondsSpinner != null) {
            timeoutSecondsSpinner.setValueFactory(
                    new SpinnerValueFactory.IntegerSpinnerValueFactory(10, 600, 120, 10));
        }
        installRankingColumns();
        refreshShellDerivedLabels();
    }

    private void installRankingColumns() {
        if (rankingTable == null || !rankingTable.getColumns().isEmpty()) {
            return;
        }
        TableColumn<BenchmarkTableRow, Integer> colRank = new TableColumn<>("順位");
        colRank.setCellValueFactory(new PropertyValueFactory<>("rank"));
        colRank.setPrefWidth(52);
        TableColumn<BenchmarkTableRow, String> colModel = new TableColumn<>("モデル");
        colModel.setCellValueFactory(new PropertyValueFactory<>("modelName"));
        colModel.setPrefWidth(200);
        TableColumn<BenchmarkTableRow, String> colAvg = new TableColumn<>("平均 ms");
        colAvg.setCellValueFactory(new PropertyValueFactory<>("avgMillis"));
        colAvg.setPrefWidth(88);
        TableColumn<BenchmarkTableRow, String> colMin = new TableColumn<>("最小 ms");
        colMin.setCellValueFactory(new PropertyValueFactory<>("minMillis"));
        colMin.setPrefWidth(88);
        TableColumn<BenchmarkTableRow, String> colMax = new TableColumn<>("最大 ms");
        colMax.setCellValueFactory(new PropertyValueFactory<>("maxMillis"));
        colMax.setPrefWidth(88);
        TableColumn<BenchmarkTableRow, Integer> colOk = new TableColumn<>("2xx 回数");
        colOk.setCellValueFactory(new PropertyValueFactory<>("successCount"));
        colOk.setPrefWidth(72);
        TableColumn<BenchmarkTableRow, String> colAtt = new TableColumn<>("実施/計画");
        colAtt.setCellValueFactory(new PropertyValueFactory<>("attempts"));
        colAtt.setPrefWidth(88);
        TableColumn<BenchmarkTableRow, String> colNotes = new TableColumn<>("備考");
        colNotes.setCellValueFactory(new PropertyValueFactory<>("notes"));
        colNotes.setPrefWidth(220);
        rankingTable.getColumns().addAll(colRank, colModel, colAvg, colMin, colMax, colOk, colAtt, colNotes);
    }

    void bindShell(MainShellController mainShell) {
        this.shell = mainShell;
        refreshShellDerivedLabels();
    }

    /** 認証パス・API モデル優先表示を、現在のシェル環境で更新する（タブ選択時にも呼ぶ）。 */
    public void refreshShellDerivedLabels() {
        refreshCredentialsPathLabel();
        refreshModelPriorityLabel();
    }

    private void refreshCredentialsPathLabel() {
        if (credentialsPathLabel == null) {
            return;
        }
        if (shell == null) {
            credentialsPathLabel.setText("（シェル未接続）");
            return;
        }
        Path p = resolveGeminiCredentialsPath(shell.snapshotUiEnv());
        credentialsPathLabel.setText("GEMINI_CREDENTIALS_JSON → " + p);
    }

    private void refreshModelPriorityLabel() {
        if (modelPriorityLabel == null) {
            return;
        }
        String defaultsJoined =
                String.join(" → ", GeminiDispatchModelTryOrderDefaults.PLANNING_CORE_FALLBACK_TRY_ORDER);
        String header =
                "【planning_core】Gemini 試行モデル列（_gemini_generate_content_with_retry）\n"
                        + "決定ルール: 呼び出し引数で model を渡したときはその1件のみ。"
                        + " 未指定のときは (1) 環境変数 GEMINI_MODEL が非空ならその1件 (2) GEMINI_MODEL_TRY_ORDER が非空ならカンマ区切りを左から順"
                        + " (3) コード既定: "
                        + defaultsJoined
                        + "\n";
        if (shell == null) {
            modelPriorityLabel.setText(header + "現在の列: （シェル未接続）");
            return;
        }
        Map<String, String> env = shell.snapshotUiEnv();
        String pinned = firstNonBlankEnv(env, "GEMINI_MODEL");
        String tryRaw = env != null ? Objects.toString(env.get("GEMINI_MODEL_TRY_ORDER"), "") : "";
        String tryParsed = parseCommaSeparatedModelOrder(tryRaw);
        String effective;
        String sourceLine;
        if (pinned != null) {
            effective = pinned;
            sourceLine = "ソース: GEMINI_MODEL（単一固定）";
        } else if (tryParsed != null) {
            effective = tryParsed;
            sourceLine = "ソース: GEMINI_MODEL_TRY_ORDER";
        } else {
            effective = defaultsJoined;
            sourceLine = "ソース: planning_core コード既定（GEMINI_MODEL_IDS_BY_QUALITY）";
        }
        modelPriorityLabel.setText(
                header
                        + "現在の列（環境タブのスナップショット）: "
                        + effective
                        + "\n"
                        + sourceLine
                        + "\n"
                        + "※ 本タブの「全モデル一括」は ComboBox 一覧の順であり、上記とは別です。");
    }

    private static String firstNonBlankEnv(Map<String, String> env, String key) {
        if (env == null || key == null) {
            return null;
        }
        String v = env.get(key);
        if (v == null) {
            return null;
        }
        String s = v.strip();
        return s.isEmpty() ? null : s;
    }

    /** カンマ区切りを「 → 」で連結。空・空白のみなら {@code null}。 */
    private static String parseCommaSeparatedModelOrder(String raw) {
        if (raw == null || raw.isBlank()) {
            return null;
        }
        List<String> parts = new ArrayList<>();
        for (String p : raw.split(",")) {
            String t = p != null ? p.strip() : "";
            if (!t.isEmpty()) {
                parts.add(t);
            }
        }
        if (parts.isEmpty()) {
            return null;
        }
        return String.join(" → ", parts);
    }

    @FXML
    private void onClearLog() {
        if (logArea != null) {
            logArea.clear();
        }
        if (summaryLabel != null) {
            summaryLabel.setText("");
        }
        if (rankingTable != null) {
            rankingTable.setItems(FXCollections.observableArrayList());
        }
    }

    @FXML
    private void onCancelBench() {
        if (worker.get() == null || !worker.get().isAlive()) {
            return;
        }
        benchmarkCancelRequested.set(true);
        appendLogLine("【中断】現在の HTTP リクエスト完了後に打ち切ります…");
    }

    @FXML
    private void onRunBenchmark() {
        if (shell == null) {
            alert(AlertType.WARNING, "メインシェルが未接続のため実行できません。");
            return;
        }
        if (worker.get() != null && worker.get().isAlive()) {
            alert(AlertType.INFORMATION, "既にベンチマークが実行中です。");
            return;
        }

        String apiKey = loadApiKeyOrAlert();
        if (apiKey == null) {
            return;
        }

        String modelRaw = modelCombo != null ? modelCombo.getEditor().getText() : "";
        String model = GeminiGenerateContentRestClient.normalizeModelId(modelRaw);
        if (model.isEmpty()) {
            alert(AlertType.WARNING, "モデル ID を入力してください。");
            return;
        }
        if (!GeminiGenerateContentRestClient.isAllowedModelId(modelRaw)) {
            alert(AlertType.WARNING, "モデル ID は英数字・ドット・ハイフン・アンダースコアのみ使用してください。");
            return;
        }

        int runs = runsSpinner != null ? runsSpinner.getValue() : 1;
        int maxTok = maxTokensSpinner != null ? maxTokensSpinner.getValue() : 64;
        int timeoutSec = timeoutSecondsSpinner != null ? timeoutSecondsSpinner.getValue() : 120;
        boolean warmup = warmupCheck != null && warmupCheck.isSelected();
        boolean stopOnTerminal = stopOnTerminalHttpCheck == null || stopOnTerminalHttpCheck.isSelected();
        String prompt = promptArea != null ? promptArea.getText() : "";
        Duration timeout = Duration.ofSeconds(Math.max(1, timeoutSec));

        beginBenchRunningUi("単体: " + model);

        Thread t =
                new Thread(
                        () -> {
                            try {
                                ModelBenchOutcome outcome =
                                        runBenchmarkOnWorker(
                                                apiKey,
                                                model,
                                                runs,
                                                maxTok,
                                                timeout,
                                                warmup,
                                                stopOnTerminal,
                                                prompt,
                                                benchmarkCancelRequested);
                                final ModelBenchOutcome finOutcome = outcome;
                                final boolean finInterrupted = outcome.userInterrupted();
                                Platform.runLater(
                                        () -> {
                                            if (summaryLabel != null) {
                                                summaryLabel.setText(finOutcome.toSummaryLine());
                                            }
                                            refreshRankingTable(List.of(finOutcome));
                                            notifyBenchmarkAlert(
                                                    AlertType.INFORMATION,
                                                    finInterrupted ? "ベンチマーク中断" : "ベンチマーク完了",
                                                    finInterrupted
                                                            ? "単体ベンチマークが中断されました。\n"
                                                                    + finOutcome.toSummaryLine()
                                                            : "単体ベンチマークが終了しました。\n"
                                                                    + finOutcome.toSummaryLine());
                                        });
                            } catch (Exception ex) {
                                Platform.runLater(
                                        () ->
                                                notifyBenchmarkAlert(
                                                        AlertType.ERROR,
                                                        "ベンチマークエラー",
                                                        "例外: "
                                                                + ex.getClass().getSimpleName()
                                                                + "\n"
                                                                + ex.getMessage()));
                            } finally {
                                endBenchRunningUi();
                            }
                        },
                        "api-model-benchmark");
        worker.set(t);
        t.setDaemon(true);
        t.start();
    }

    @FXML
    private void onRunAllModelsBenchmark() {
        if (shell == null) {
            alert(AlertType.WARNING, "メインシェルが未接続のため実行できません。");
            return;
        }
        if (worker.get() != null && worker.get().isAlive()) {
            alert(AlertType.INFORMATION, "既にベンチマークが実行中です。");
            return;
        }

        String apiKey = loadApiKeyOrAlert();
        if (apiKey == null) {
            return;
        }

        List<String> models = orderedAllowedModelsFromCombo();
        if (models.isEmpty()) {
            alert(
                    AlertType.WARNING,
                    "ComboBox の一覧に有効なモデル ID がありません。\n"
                            + "英数字・ドット・ハイフン・アンダースコアのみの ID をリストに追加してください。");
            return;
        }

        int runs = runsSpinner != null ? runsSpinner.getValue() : 1;
        int maxTok = maxTokensSpinner != null ? maxTokensSpinner.getValue() : 64;
        int timeoutSec = timeoutSecondsSpinner != null ? timeoutSecondsSpinner.getValue() : 120;
        boolean warmup = warmupCheck != null && warmupCheck.isSelected();
        boolean stopOnTerminal = stopOnTerminalHttpCheck == null || stopOnTerminalHttpCheck.isSelected();
        String prompt = promptArea != null ? promptArea.getText() : "";
        Duration timeout = Duration.ofSeconds(Math.max(1, timeoutSec));

        appendLogLine(
                "=== 全モデル一括: "
                        + models.size()
                        + " 件（試行回数="
                        + runs
                        + ", maxOutputTokens="
                        + maxTok
                        + ", 404/429打ち切り="
                        + stopOnTerminal
                        + "）===");

        beginBenchRunningUi("全モデル: " + models.size() + " 件");

        Thread t =
                new Thread(
                        () -> {
                            boolean anyInterrupted = false;
                            try {
                                List<ModelBenchOutcome> outcomes = new ArrayList<>();
                                for (String model : models) {
                                    if (benchmarkCancelRequested.get()) {
                                        appendLogLine("=== 全モデル一括 中断（未実行モデルあり）===");
                                        anyInterrupted = true;
                                        break;
                                    }
                                    ModelBenchOutcome o =
                                            runBenchmarkOnWorker(
                                                    apiKey,
                                                    model,
                                                    runs,
                                                    maxTok,
                                                    timeout,
                                                    warmup,
                                                    stopOnTerminal,
                                                    prompt,
                                                    benchmarkCancelRequested);
                                    outcomes.add(o);
                                    if (o.userInterrupted()) {
                                        anyInterrupted = true;
                                        appendLogLine("=== 全モデル一括 中断 ===");
                                        break;
                                    }
                                }
                                StringBuilder sb = new StringBuilder();
                                sb.append("【全モデル集計】\n");
                                for (ModelBenchOutcome o : outcomes) {
                                    sb.append(o.toSummaryLine()).append('\n');
                                }
                                String text = sb.toString().strip();
                                boolean finInterrupted = anyInterrupted;
                                List<ModelBenchOutcome> finOutcomes = List.copyOf(outcomes);
                                Platform.runLater(
                                        () -> {
                                            if (summaryLabel != null) {
                                                summaryLabel.setText(text);
                                            }
                                            refreshRankingTable(finOutcomes);
                                            String body;
                                            if (finOutcomes.isEmpty() && finInterrupted) {
                                                body = "中断により、まだどのモデルも実行されていません。";
                                            } else if (finInterrupted) {
                                                body =
                                                        "一部のみ実行されました（"
                                                                + finOutcomes.size()
                                                                + " モデル）。\nランキング表を確認してください。";
                                            } else {
                                                body =
                                                        finOutcomes.size()
                                                                + " モデルのベンチマークが終了しました。\nランキング表を確認してください。";
                                            }
                                            notifyBenchmarkAlert(
                                                    AlertType.INFORMATION,
                                                    finInterrupted ? "全モデル一括 中断" : "全モデル一括 完了",
                                                    body);
                                        });
                                if (!finInterrupted) {
                                    appendLogLine("=== 全モデル一括 完了 ===");
                                }
                            } catch (Exception ex) {
                                Platform.runLater(
                                        () ->
                                                notifyBenchmarkAlert(
                                                        AlertType.ERROR,
                                                        "ベンチマークエラー",
                                                        "例外: "
                                                                + ex.getClass().getSimpleName()
                                                                + "\n"
                                                                + ex.getMessage()));
                            } finally {
                                endBenchRunningUi();
                            }
                        },
                        "api-model-benchmark-all");
        worker.set(t);
        t.setDaemon(true);
        t.start();
    }

    private void beginBenchRunningUi(String runningHint) {
        benchmarkCancelRequested.set(false);
        Platform.runLater(
                () -> {
                    if (benchProgressBar != null) {
                        benchProgressBar.setManaged(true);
                        benchProgressBar.setVisible(true);
                        benchProgressBar.setProgress(ProgressBar.INDETERMINATE_PROGRESS);
                    }
                    if (benchRunningLabel != null) {
                        benchRunningLabel.setManaged(true);
                        benchRunningLabel.setVisible(true);
                        benchRunningLabel.setText(runningHint != null ? runningHint : "実行中…");
                    }
                    if (cancelBenchButton != null) {
                        cancelBenchButton.setDisable(false);
                    }
                    setBenchInputsDisabled(true);
                    if (runButton != null) {
                        runButton.setDisable(true);
                    }
                    if (runAllModelsButton != null) {
                        runAllModelsButton.setDisable(true);
                    }
                    if (clearButton != null) {
                        clearButton.setDisable(true);
                    }
                });
    }

    private void endBenchRunningUi() {
        benchmarkCancelRequested.set(false);
        Platform.runLater(
                () -> {
                    if (benchProgressBar != null) {
                        benchProgressBar.setProgress(0);
                        benchProgressBar.setVisible(false);
                        benchProgressBar.setManaged(false);
                    }
                    if (benchRunningLabel != null) {
                        benchRunningLabel.setText("");
                        benchRunningLabel.setVisible(false);
                        benchRunningLabel.setManaged(false);
                    }
                    if (cancelBenchButton != null) {
                        cancelBenchButton.setDisable(true);
                    }
                    setBenchInputsDisabled(false);
                    if (runButton != null) {
                        runButton.setDisable(false);
                    }
                    if (runAllModelsButton != null) {
                        runAllModelsButton.setDisable(false);
                    }
                    if (clearButton != null) {
                        clearButton.setDisable(false);
                    }
                    worker.set(null);
                });
    }

    private void setBenchInputsDisabled(boolean disabled) {
        if (modelCombo != null) {
            modelCombo.setDisable(disabled);
        }
        if (runsSpinner != null) {
            runsSpinner.setDisable(disabled);
        }
        if (maxTokensSpinner != null) {
            maxTokensSpinner.setDisable(disabled);
        }
        if (timeoutSecondsSpinner != null) {
            timeoutSecondsSpinner.setDisable(disabled);
        }
        if (warmupCheck != null) {
            warmupCheck.setDisable(disabled);
        }
        if (stopOnTerminalHttpCheck != null) {
            stopOnTerminalHttpCheck.setDisable(disabled);
        }
        if (completionNotifyCheck != null) {
            completionNotifyCheck.setDisable(disabled);
        }
        if (promptArea != null) {
            promptArea.setDisable(disabled);
        }
    }

    private void refreshRankingTable(List<ModelBenchOutcome> outcomes) {
        if (rankingTable == null || outcomes == null || outcomes.isEmpty()) {
            if (rankingTable != null) {
                rankingTable.setItems(FXCollections.observableArrayList());
            }
            return;
        }
        List<ModelBenchOutcome> sorted = new ArrayList<>(outcomes);
        Comparator<ModelBenchOutcome> cmp =
                Comparator.comparing((ModelBenchOutcome o) -> o.http2xxCount() <= 0)
                        .thenComparing(
                                o ->
                                        o.http2xxCount() > 0 && !Double.isNaN(o.avgMs())
                                                ? o.avgMs()
                                                : Double.POSITIVE_INFINITY)
                        .thenComparing(ModelBenchOutcome::model);
        sorted.sort(cmp);
        ObservableList<BenchmarkTableRow> rows = FXCollections.observableArrayList();
        int rank = 1;
        for (ModelBenchOutcome o : sorted) {
            rows.add(BenchmarkTableRow.ranked(rank++, o));
        }
        rankingTable.setItems(rows);
    }

    private void notifyBenchmarkAlert(AlertType type, String title, String message) {
        if (completionNotifyCheck == null || !completionNotifyCheck.isSelected()) {
            return;
        }
        Platform.runLater(
                () -> {
                    Alert a = new Alert(type != null ? type : AlertType.INFORMATION, message);
                    a.setTitle("APIモデルベンチマーク");
                    a.setHeaderText(title);
                    a.show();
                });
    }

    /**
     * ComboBox の {@code items} を上から順に走査し、有効なモデル ID を重複なく返す（編集行のテキストは含めない）。
     */
    private List<String> orderedAllowedModelsFromCombo() {
        LinkedHashSet<String> out = new LinkedHashSet<>();
        if (modelCombo != null && modelCombo.getItems() != null) {
            for (String item : modelCombo.getItems()) {
                if (GeminiGenerateContentRestClient.isAllowedModelId(item)) {
                    out.add(GeminiGenerateContentRestClient.normalizeModelId(item));
                }
            }
        }
        return List.copyOf(out);
    }

    /** 認証ファイルの読込・復号。失敗時はアラートのみで {@code null}。 */
    private String loadApiKeyOrAlert() {
        Path credPath = resolveGeminiCredentialsPath(shell.snapshotUiEnv());
        if (!Files.isRegularFile(credPath)) {
            alert(
                    AlertType.WARNING,
                    "認証ファイルが見つかりません。\n"
                            + credPath
                            + "\n環境変数タブで GEMINI_CREDENTIALS_JSON を設定するか、既定パスにファイルを置いてください。");
            return null;
        }
        try {
            String json = Files.readString(credPath);
            String apiKey =
                    GeminiCredentialsV2Crypto.decryptGeminiApiKeyFromJsonString(
                            json, GeminiCredentialsV2Crypto.DEFAULT_PASSPHRASE);
            if (apiKey == null || apiKey.isBlank()) {
                alert(AlertType.ERROR, "復号した API キーが空です。");
                return null;
            }
            return apiKey;
        } catch (Exception ex) {
            alert(AlertType.ERROR, "認証 JSON の読込・復号に失敗しました。\n" + ex.getMessage());
            return null;
        }
    }

    private static boolean shouldCutRemainingTrials(int httpStatus) {
        return httpStatus == 404 || httpStatus == 429;
    }

    private ModelBenchOutcome runBenchmarkOnWorker(
            String apiKey,
            String model,
            int runs,
            int maxTok,
            Duration timeout,
            boolean warmup,
            boolean stopOnTerminalHttp,
            String prompt,
            AtomicBoolean cancelRequested) {
        if (cancelRequested.get()) {
            appendLogLine("--- 中断のためスキップ model=" + model + " ---");
            return new ModelBenchOutcome(
                    model, 0, 0, runs, Double.NaN, Double.NaN, Double.NaN, "ユーザー中断", true);
        }

        appendLogLine("--- 開始 model=" + model + " runs=" + runs + " maxOutputTokens=" + maxTok + " ---");
        if (warmup) {
            if (cancelRequested.get()) {
                return interruptedOutcome(model, runs, 0, 0, Double.NaN, Double.NaN, Double.NaN, "");
            }
            try {
                appendLogLine("[ウォームアップ] 送信中…");
                CallResult w =
                        GeminiGenerateContentRestClient.generateContent(
                                apiKey, model, prompt, maxTok, timeout);
                appendLogLine(
                        String.format(
                                Locale.ROOT,
                                "[ウォームアップ] HTTP %d  %.2f ms",
                                w.httpStatus(),
                                w.wallTimeMs()));
                if (stopOnTerminalHttp && shouldCutRemainingTrials(w.httpStatus())) {
                    logHttpFailureDetailOnce(w);
                    appendLogLine("[ウォームアップ] 計測ループはスキップ（HTTP " + w.httpStatus() + " で打ち切り）。");
                    appendLogLine("--- 終了 model=" + model + " ---");
                    return new ModelBenchOutcome(
                            model,
                            0,
                            0,
                            runs,
                            Double.NaN,
                            Double.NaN,
                            Double.NaN,
                            "HTTP " + w.httpStatus() + "（ウォームアップで打ち切り）",
                            false);
                }
                if (!w.errorSummary().isEmpty()) {
                    appendLogLine("[ウォームアップ] " + w.errorSummary());
                }
            } catch (Exception ex) {
                appendLogLine("[ウォームアップ] 失敗: " + ex.getMessage());
            }
        }

        if (cancelRequested.get()) {
            return interruptedOutcome(model, runs, 0, 0, Double.NaN, Double.NaN, Double.NaN, "");
        }

        double sumMs = 0;
        int counted = 0;
        double minMs = Double.POSITIVE_INFINITY;
        double maxMs = 0;
        boolean failureDetailLogged = false;
        String supplement = "";

        for (int i = 1; i <= runs; i++) {
            if (cancelRequested.get()) {
                appendLogLine("--- ユーザー中断 model=" + model + " ---");
                appendLogLine("--- 終了 model=" + model + " ---");
                String sup = supplement.isBlank() ? "" : supplement;
                if (counted > 0) {
                    return new ModelBenchOutcome(
                            model, counted, i - 1, runs, sumMs / counted, minMs, maxMs, sup, true);
                }
                return new ModelBenchOutcome(
                        model, 0, i - 1, runs, Double.NaN, Double.NaN, Double.NaN, sup, true);
            }
            try {
                CallResult r =
                        GeminiGenerateContentRestClient.generateContent(
                                apiKey, model, prompt, maxTok, timeout);
                double ms = r.wallTimeMs();
                if (r.httpStatus() >= 200 && r.httpStatus() < 300) {
                    sumMs += ms;
                    counted++;
                    minMs = Math.min(minMs, ms);
                    maxMs = Math.max(maxMs, ms);
                }
                String line =
                        String.format(Locale.ROOT, "#%d  HTTP %d  %.2f ms", i, r.httpStatus(), ms);
                appendLogLine(line);
                if (stopOnTerminalHttp && shouldCutRemainingTrials(r.httpStatus())) {
                    if (!failureDetailLogged) {
                        logHttpFailureDetailOnce(r);
                        failureDetailLogged = true;
                    }
                    int remaining = runs - i;
                    appendLogLine(
                            "  （HTTP "
                                    + r.httpStatus()
                                    + " のため残り "
                                    + remaining
                                    + " 回を省略）");
                    supplement = "打ち切り HTTP " + r.httpStatus();
                    appendLogLine("--- 終了 model=" + model + " ---");
                    if (counted > 0) {
                        return new ModelBenchOutcome(
                                model,
                                counted,
                                i,
                                runs,
                                sumMs / counted,
                                minMs,
                                maxMs,
                                supplement,
                                false);
                    }
                    return new ModelBenchOutcome(
                            model,
                            0,
                            i,
                            runs,
                            Double.NaN,
                            Double.NaN,
                            Double.NaN,
                            supplement,
                            false);
                }
                if (r.httpStatus() < 200 || r.httpStatus() >= 300) {
                    if (!failureDetailLogged) {
                        logHttpFailureDetailOnce(r);
                        failureDetailLogged = true;
                    } else {
                        appendLogLine("  （詳細ログは先頭の失敗応答を参照）");
                    }
                }
            } catch (Exception ex) {
                appendLogLine("#" + i + "  例外: " + ex.getClass().getSimpleName() + " — " + ex.getMessage());
            }
        }

        appendLogLine("--- 終了 model=" + model + " ---");

        if (counted > 0) {
            return new ModelBenchOutcome(
                    model, counted, runs, runs, sumMs / counted, minMs, maxMs, supplement, false);
        }
        return new ModelBenchOutcome(
                model, 0, runs, runs, Double.NaN, Double.NaN, Double.NaN, supplement, false);
    }

    private static ModelBenchOutcome interruptedOutcome(
            String model,
            int plannedRuns,
            int http2xx,
            int measured,
            double avg,
            double min,
            double max,
            String supplement) {
        return new ModelBenchOutcome(
                model,
                http2xx,
                measured,
                plannedRuns,
                avg,
                min,
                max,
                supplement != null ? supplement : "",
                true);
    }

    private void logHttpFailureDetailOnce(CallResult r) {
        if (r.bodyPreview() != null && !r.bodyPreview().isEmpty()) {
            String one = r.bodyPreview().strip();
            if (one.length() > 360) {
                one = one.substring(0, 360) + "…";
            }
            appendLogLine("  body: " + one);
        }
        if (r.errorSummary() != null && !r.errorSummary().isEmpty()) {
            String e = r.errorSummary().strip();
            if (e.length() > 320) {
                e = e.substring(0, 320) + "…";
            }
            appendLogLine("  err: " + e);
        }
    }

    private void appendLogLine(String line) {
        Platform.runLater(
                () -> {
                    if (logArea == null) {
                        return;
                    }
                    logArea.appendText(line + "\n");
                });
    }

    private static void alert(AlertType type, String msg) {
        Platform.runLater(
                () -> {
                    Alert a = new Alert(type, msg);
                    a.setHeaderText(null);
                    a.show();
                });
    }

    /**
     * {@link EnvTabController} の解決と同じ: 環境マップの {@link AppPaths#KEY_GEMINI_CREDENTIALS_JSON} が非空ならそのパス、
     * 空のときは {@code resolveRepoRoot(ui)/code/gemini_credentials.encrypted.json}。
     */
    private static Path resolveGeminiCredentialsPath(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String raw = u.get(AppPaths.KEY_GEMINI_CREDENTIALS_JSON);
        if (raw != null && !raw.isBlank()) {
            return Path.of(raw.strip()).toAbsolutePath().normalize();
        }
        return AppPaths.resolveRepoRoot(u)
                .resolve("code")
                .resolve("gemini_credentials.encrypted.json")
                .toAbsolutePath()
                .normalize();
    }
}
