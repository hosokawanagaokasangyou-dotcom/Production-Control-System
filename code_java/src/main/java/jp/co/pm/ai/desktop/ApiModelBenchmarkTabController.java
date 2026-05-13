package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.concurrent.atomic.AtomicReference;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.TextArea;

import jp.co.pm.ai.desktop.benchmark.GeminiGenerateContentRestClient;
import jp.co.pm.ai.desktop.benchmark.GeminiGenerateContentRestClient.CallResult;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.crypto.GeminiCredentialsV2Crypto;

/**
 * Gemini {@code generateContent} の往復レイテンシを計測するメインシェルタブ。
 */
public final class ApiModelBenchmarkTabController {

    private static final List<String> DEFAULT_MODELS =
            List.of(
                    "gemini-2.5-flash",
                    "gemini-2.5-flash-lite",
                    "gemini-2.0-flash",
                    "gemini-2.0-flash-lite");

    private MainShellController shell;

    @FXML
    private Label credentialsPathLabel;

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
    private Label summaryLabel;

    private final AtomicReference<Thread> worker = new AtomicReference<>();

    private record ModelBenchOutcome(
            String model,
            int http2xxCount,
            int measuredAttempts,
            int plannedRuns,
            double avgMs,
            double minMs,
            double maxMs,
            String supplement) {

        String toSummaryLine() {
            if (http2xxCount <= 0) {
                String s =
                        model
                                + ": HTTP 2xx なし（実施 "
                                + measuredAttempts
                                + "/"
                                + plannedRuns
                                + "）";
                if (supplement != null && !supplement.isBlank()) {
                    s += " — " + supplement.strip();
                }
                return s;
            }
            String s =
                    String.format(
                            Locale.ROOT,
                            "%s: 平均 %.2f ms / 最小 %.2f / 最大 %.2f（n=%d）",
                            model,
                            avgMs,
                            minMs,
                            maxMs,
                            http2xxCount);
            if (supplement != null && !supplement.isBlank()) {
                s += " — " + supplement.strip();
            }
            return s;
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
        refreshCredentialsPathLabel();
    }

    void bindShell(MainShellController mainShell) {
        this.shell = mainShell;
        refreshCredentialsPathLabel();
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

    @FXML
    private void onClearLog() {
        if (logArea != null) {
            logArea.clear();
        }
        if (summaryLabel != null) {
            summaryLabel.setText("");
        }
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

        setBenchButtonsDisabled(true);

        Thread t =
                new Thread(
                        () -> {
                            try {
                                ModelBenchOutcome outcome =
                                        runBenchmarkOnWorker(
                                                apiKey, model, runs, maxTok, timeout, warmup, stopOnTerminal, prompt);
                                String line = outcome.toSummaryLine();
                                Platform.runLater(
                                        () -> {
                                            if (summaryLabel != null) {
                                                summaryLabel.setText(line);
                                            }
                                        });
                            } finally {
                                finishBenchWorker();
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

        setBenchButtonsDisabled(true);

        Thread t =
                new Thread(
                        () -> {
                            try {
                                List<ModelBenchOutcome> outcomes = new ArrayList<>();
                                for (String model : models) {
                                    ModelBenchOutcome o =
                                            runBenchmarkOnWorker(
                                                    apiKey,
                                                    model,
                                                    runs,
                                                    maxTok,
                                                    timeout,
                                                    warmup,
                                                    stopOnTerminal,
                                                    prompt);
                                    outcomes.add(o);
                                }
                                StringBuilder sb = new StringBuilder();
                                sb.append("【全モデル集計】\n");
                                for (ModelBenchOutcome o : outcomes) {
                                    sb.append(o.toSummaryLine()).append('\n');
                                }
                                String text = sb.toString().strip();
                                Platform.runLater(
                                        () -> {
                                            if (summaryLabel != null) {
                                                summaryLabel.setText(text);
                                            }
                                        });
                                appendLogLine("=== 全モデル一括 完了 ===");
                            } finally {
                                finishBenchWorker();
                            }
                        },
                        "api-model-benchmark-all");
        worker.set(t);
        t.setDaemon(true);
        t.start();
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

    private void setBenchButtonsDisabled(boolean disabled) {
        Platform.runLater(
                () -> {
                    if (runButton != null) {
                        runButton.setDisable(disabled);
                    }
                    if (runAllModelsButton != null) {
                        runAllModelsButton.setDisable(disabled);
                    }
                    if (clearButton != null) {
                        clearButton.setDisable(disabled);
                    }
                });
    }

    private void finishBenchWorker() {
        Platform.runLater(
                () -> {
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
            String prompt) {
        appendLogLine("--- 開始 model=" + model + " runs=" + runs + " maxOutputTokens=" + maxTok + " ---");
        if (warmup) {
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
                            "HTTP " + w.httpStatus() + "（ウォームアップで打ち切り）");
                }
                if (!w.errorSummary().isEmpty()) {
                    appendLogLine("[ウォームアップ] " + w.errorSummary());
                }
            } catch (Exception ex) {
                appendLogLine("[ウォームアップ] 失敗: " + ex.getMessage());
            }
        }

        double sumMs = 0;
        int counted = 0;
        double minMs = Double.POSITIVE_INFINITY;
        double maxMs = 0;
        boolean failureDetailLogged = false;
        String supplement = "";

        for (int i = 1; i <= runs; i++) {
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
                                supplement);
                    }
                    return new ModelBenchOutcome(
                            model,
                            0,
                            i,
                            runs,
                            Double.NaN,
                            Double.NaN,
                            Double.NaN,
                            supplement);
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
                    model, counted, runs, runs, sumMs / counted, minMs, maxMs, supplement);
        }
        return new ModelBenchOutcome(
                model, 0, runs, runs, Double.NaN, Double.NaN, Double.NaN, supplement);
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
