package jp.co.pm.ai.desktop.dispatch;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import java.util.concurrent.TimeUnit;
import java.util.function.Consumer;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;

/** Runs {@code attendance_overtime_preview.py} and returns parsed preview JSON. */
public final class AttendanceOvertimePreviewPython {

    private AttendanceOvertimePreviewPython() {}

    public static AttendanceOvertimePreview.Preview load(
            Path pythonExe,
            Path pythonScriptDir,
            Map<String, String> extraUiEnv,
            Consumer<String> logLine)
            throws Exception {
        Path script = pythonScriptDir.resolve("attendance_overtime_preview.py");
        if (!Files.isRegularFile(script)) {
            throw new IllegalStateException(
                    "missing attendance_overtime_preview.py in " + pythonScriptDir);
        }
        ProcessBuilder pb =
                new ProcessBuilder(
                        pythonExe.toString(), script.toAbsolutePath().toString());
        pb.directory(pythonScriptDir.toFile());
        pb.redirectErrorStream(true);
        PythonProcessRunner.mergeUiEnvIntoProcess(pb, extraUiEnv, pythonScriptDir);
        Process p = pb.start();
        StringBuilder merged = new StringBuilder();
        try (BufferedReader br =
                new BufferedReader(
                        new InputStreamReader(p.getInputStream(), StandardCharsets.UTF_8))) {
            String line;
            while ((line = br.readLine()) != null) {
                if (logLine != null) {
                    logLine.accept(line);
                }
                if (merged.length() > 0) {
                    merged.append('\n');
                }
                merged.append(line);
            }
        }
        boolean finished = p.waitFor(600, TimeUnit.SECONDS);
        if (!finished) {
            p.destroyForcibly();
            throw new IllegalStateException("attendance_overtime_preview.py timeout");
        }
        if (p.exitValue() != 0) {
            throw new IllegalStateException(
                    "attendance_overtime_preview.py exit "
                            + p.exitValue()
                            + ": "
                            + merged);
        }
        AttendanceOvertimePreview.Preview preview =
                AttendanceOvertimePreview.parseJson(merged.toString());
        if (!preview.ok()) {
            throw new IllegalStateException(
                    preview.error() != null && !preview.error().isBlank()
                            ? preview.error()
                            : "勤怠プレビューの取得に失敗しました");
        }
        return preview;
    }
}
