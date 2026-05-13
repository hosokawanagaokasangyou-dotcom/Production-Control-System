package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
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
                    "gemini-2.0-flash",
                    "gemini-2.0-flash-lite",
                    "gemini-1.5-flash",
                    "gemini-1.5-flash-8b",
                    "gemini-2.5-flash-preview-05-20");

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
    private Button runButton;

    @FXML
    private Button clearButton;

    @FXML
    private TextArea promptArea;

    @FXML
    private TextArea logArea;

    @FXML
    private Label summaryLabel;

    private final AtomicReference<Thread> worker = new AtomicReference<>();

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

        Path credPath = resolveGeminiCredentialsPath(shell.snapshotUiEnv());
        if (!Files.isRegularFile(credPath)) {
            alert(
                    AlertType.WARNING,
                    "認証ファイルが見つかりません。\n"
                            + credPath
                            + "\n環境変数タブで GEMINI_CREDENTIALS_JSON を設定するか、既定パスにファイルを置いてください。");
            return;
        }

        final String apiKey;
        try {
            String json = Files.readString(credPath);
            apiKey =
                    GeminiCredentialsV2Crypto.decryptGeminiApiKeyFromJsonString(
                            json, GeminiCredentialsV2Crypto.DEFAULT_PASSPHRASE);
        } catch (Exception ex) {
            alert(AlertType.ERROR, "認証 JSON の読込・復号に失敗しました。\n" + ex.getMessage());
            return;
        }
        if (apiKey == null || apiKey.isBlank()) {
            alert(AlertType.ERROR, "復号した API キーが空です。");
            return;
        }

        String modelRaw = modelCombo != null ? modelCombo.getEditor().getText() : "";
        String model = GeminiGenerateContentRestClient.normalizeModelId(modelRaw);
        if (model.isEmpty()) {
            alert(AlertType.WARNING, "モデル ID を入力してください。");
            return;
        }

        int runs = runsSpinner != null ? runsSpinner.getValue() : 1;
        int maxTok = maxTokensSpinner != null ? maxTokensSpinner.getValue() : 64;
        int timeoutSec = timeoutSecondsSpinner != null ? timeoutSecondsSpinner.getValue() : 120;
        boolean warmup = warmupCheck != null && warmupCheck.isSelected();
        String prompt = promptArea != null ? promptArea.getText() : "";
        Duration timeout = Duration.ofSeconds(Math.max(1, timeoutSec));

        if (runButton != null) {
            runButton.setDisable(true);
        }
        if (clearButton != null) {
            clearButton.setDisable(true);
        }

        Thread t =
                new Thread(
                        () -> {
                            try {
                                runBenchmarkOnWorker(apiKey, model, runs, maxTok, timeout, warmup, prompt);
                            } finally {
                                Platform.runLater(
                                        () -> {
                                            if (runButton != null) {
                                                runButton.setDisable(false);
                                            }
                                            if (clearButton != null) {
                                                clearButton.setDisable(false);
                                            }
                                            worker.set(null);
                                        });
                            }
                        },
                        "api-model-benchmark");
        worker.set(t);
        t.setDaemon(true);
        t.start();
    }

    private void runBenchmarkOnWorker(
            String apiKey,
            String model,
            int runs,
            int maxTok,
            Duration timeout,
            boolean warmup,
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
                        String.format(
                                Locale.ROOT,
                                "#%d  HTTP %d  %.2f ms",
                                i, r.httpStatus(), ms);
                appendLogLine(line);
                if (!r.bodyPreview().isEmpty()) {
                    appendLogLine("  body: " + r.bodyPreview());
                }
                if (!r.errorSummary().isEmpty()) {
                    appendLogLine("  err: " + r.errorSummary());
                }
            } catch (Exception ex) {
                appendLogLine("#" + i + "  例外: " + ex.getClass().getSimpleName() + " — " + ex.getMessage());
            }
        }

        String summary;
        if (counted > 0) {
            double avg = sumMs / counted;
            summary =
                    String.format(
                            Locale.ROOT,
                            "集計（HTTP 2xx のみ）: 平均 %.2f ms / 最小 %.2f ms / 最大 %.2f ms（n=%d）",
                            avg, minMs, maxMs, counted);
        } else {
            summary = "集計: 成功レスポンスが無かったため平均を算出できませんでした。";
        }
        String finalSummary = summary;
        Platform.runLater(
                () -> {
                    appendLogLine("--- 終了 ---");
                    if (summaryLabel != null) {
                        summaryLabel.setText(finalSummary);
                    }
                });
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
