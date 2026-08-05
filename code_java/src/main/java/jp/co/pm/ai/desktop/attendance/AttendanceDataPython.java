package jp.co.pm.ai.desktop.attendance;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.TimeUnit;
import java.util.function.Consumer;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.dispatch.AttendanceOvertimePreview.MasterReadSummaryJson;

/** Runs {@code attendance_data_io.py} and returns parsed JSON. */
public final class AttendanceDataPython {

    private static final ObjectMapper JSON = new ObjectMapper();

    private AttendanceDataPython() {}

    public static JsonNode run(
            Path pythonExe,
            Path pythonScriptDir,
            Map<String, String> extraUiEnv,
            Consumer<String> logLine,
            String... args)
            throws Exception {
        Path script = pythonScriptDir.resolve("attendance_data_io.py");
        if (!Files.isRegularFile(script)) {
            throw new IllegalStateException("missing attendance_data_io.py in " + pythonScriptDir);
        }
        List<String> cmd = new ArrayList<>();
        cmd.add(pythonExe.toString());
        cmd.add(script.toAbsolutePath().toString());
        for (String a : args) {
            cmd.add(a);
        }
        ProcessBuilder pb = new ProcessBuilder(cmd);
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
            throw new IllegalStateException("attendance_data_io.py timeout");
        }
        if (p.exitValue() != 0) {
            throw new IllegalStateException(
                    "attendance_data_io.py exit " + p.exitValue() + ": " + merged);
        }
        String payload = MasterReadSummaryJson.extractLastJsonLine(merged.toString());
        return JSON.readTree(payload);
    }
}
