package jp.co.pm.ai.desktop.dispatch;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.concurrent.TimeUnit;
import java.util.function.Consumer;

import javafx.application.Platform;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.bridge.StagePythonExecutable;
import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * 段階2.5 背景: 学習アーカイブ・乖離指標・速度分布（daemon スレッド、同一 job_id はマージ）。
 */
public final class DispatchLearningArchiveBackgroundService {

    public record ArchiveJobDescriptor(
            String jobId,
            Path descriptorJson,
            Path archiveRoot,
            Path stage2Raw,
            Path afterStage25,
            Path aladdinJson,
            Path planInput,
            Path pythonExe,
            Map<String, String> uiEnv) {}

    private final Consumer<String> logSink;
    private final LinkedBlockingQueue<ArchiveJobDescriptor> queue = new LinkedBlockingQueue<>();
    private final Thread worker;

    public DispatchLearningArchiveBackgroundService(Consumer<String> logSink) {
        this.logSink = logSink != null ? logSink : s -> {};
        this.worker =
                new Thread(this::runLoop, "dispatch-learning-archive-bg");
        this.worker.setDaemon(true);
        this.worker.start();
    }

    public void enqueue(ArchiveJobDescriptor job) {
        if (job == null || job.jobId() == null || job.jobId().isBlank()) {
            return;
        }
        synchronized (queue) {
            queue.removeIf(j -> job.jobId().equals(j.jobId()));
            queue.offer(job);
        }
    }

    private void runLoop() {
        while (!Thread.currentThread().isInterrupted()) {
            try {
                ArchiveJobDescriptor job = queue.take();
                runOne(job);
            } catch (InterruptedException ex) {
                Thread.currentThread().interrupt();
                return;
            }
        }
    }

    private void runOne(ArchiveJobDescriptor job) {
        try {
            Map<String, String> ui = job.uiEnv() != null ? job.uiEnv() : Map.of();
            Path scriptDir = AppPaths.resolvePythonScriptDir(ui);
            Path script = scriptDir.resolve("dispatch_learning_archive.py");
            if (!Files.isRegularFile(script)) {
                appendLog("[bg] 学習アーカイブ失敗: dispatch_learning_archive.py が見つかりません");
                return;
            }
            Path py = resolvePythonExecutable(job);
            if (py == null || py.toString().isBlank()) {
                appendLog("[bg] 学習アーカイブ失敗: Python 実行ファイルが未設定です");
                return;
            }
            String pyToken = py.toString().trim();
            if (!isBarePythonCommand(pyToken) && !Files.isRegularFile(py)) {
                appendLog(
                        "[bg] 学習アーカイブ失敗: Python 実行ファイルが見つかりません: "
                                + py.toAbsolutePath().normalize());
                return;
            }
            java.util.ArrayList<String> cmd = new java.util.ArrayList<>();
            cmd.add(pyToken);
            cmd.add(script.toAbsolutePath().toString());
            if (job.descriptorJson() != null && Files.isRegularFile(job.descriptorJson())) {
                cmd.add("--descriptor");
                cmd.add(job.descriptorJson().toAbsolutePath().toString());
            } else {
                cmd.add("--job-id");
                cmd.add(job.jobId());
                cmd.add("--archive-root");
                cmd.add(job.archiveRoot().toString());
                if (job.stage2Raw() != null) {
                    cmd.add("--stage2-raw");
                    cmd.add(job.stage2Raw().toString());
                }
                if (job.afterStage25() != null) {
                    cmd.add("--after-stage2-5");
                    cmd.add(job.afterStage25().toString());
                }
                if (job.aladdinJson() != null) {
                    cmd.add("--aladdin-json");
                    cmd.add(job.aladdinJson().toString());
                }
                if (job.planInput() != null) {
                    cmd.add("--plan-input");
                    cmd.add(job.planInput().toString());
                }
            }
            ProcessBuilder pb = new ProcessBuilder(cmd);
            pb.directory(scriptDir.toFile());
            pb.redirectErrorStream(true);
            PythonProcessRunner.mergeUiEnvIntoProcess(pb, ui, scriptDir);
            Process p = pb.start();
            try (BufferedReader br =
                    new BufferedReader(
                            new InputStreamReader(p.getInputStream(), StandardCharsets.UTF_8))) {
                String line;
                while ((line = br.readLine()) != null) {
                    String out = line;
                    appendLog(out);
                }
            }
            boolean finished = p.waitFor(3600, TimeUnit.SECONDS);
            if (!finished) {
                p.destroyForcibly();
                appendLog("[bg] 学習アーカイブ タイムアウト job_id=" + job.jobId());
            } else if (p.exitValue() != 0) {
                appendLog("[bg] 学習アーカイブ 終了コード " + p.exitValue() + " job_id=" + job.jobId());
            }
        } catch (Exception ex) {
            appendLog(
                    "[bg] 学習アーカイブ例外 job_id="
                            + job.jobId()
                            + ": "
                            + (ex.getMessage() != null ? ex.getMessage() : ex));
        }
    }

    private void appendLog(String line) {
        Platform.runLater(() -> logSink.accept(line));
    }

    private static Path resolvePythonExecutable(ArchiveJobDescriptor job) {
        Path fromJob = job.pythonExe();
        if (fromJob != null && !fromJob.toString().isBlank()) {
            if (isBarePythonCommand(fromJob.toString().trim()) || Files.isRegularFile(fromJob)) {
                return fromJob;
            }
        }
        return StagePythonExecutable.resolve(job.uiEnv() != null ? job.uiEnv() : Map.of());
    }

    /** {@code python} / {@code python3} / {@code py} 等、PATH 解決されるコマンド名。 */
    private static boolean isBarePythonCommand(String token) {
        if (token == null || token.isBlank()) {
            return false;
        }
        if (token.contains("/") || token.contains("\\")) {
            return false;
        }
        String low = token.toLowerCase(java.util.Locale.ROOT);
        return low.equals("python")
                || low.equals("python3")
                || low.equals("py")
                || low.startsWith("py -")
                || low.startsWith("py-");
    }
}
