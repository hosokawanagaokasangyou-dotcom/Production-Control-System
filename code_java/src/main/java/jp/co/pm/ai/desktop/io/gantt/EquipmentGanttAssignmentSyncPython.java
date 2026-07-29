package jp.co.pm.ai.desktop.io.gantt;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.CompletableFuture;
import java.util.function.Consumer;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.config.AppPaths;

/** {@code sync_equipment_gantt_assignment.py} を非同期実行する。 */
public final class EquipmentGanttAssignmentSyncPython {

    private static final String SCRIPT = "sync_equipment_gantt_assignment.py";

    private EquipmentGanttAssignmentSyncPython() {}

    public record Request(
            Path contractPath,
            Path updatesJsonPath,
            Path planXlsxPath,
            String expectedTimelineHash,
            String confirmToken,
            boolean forceWarnings,
            boolean dryRun) {}

    public static CompletableFuture<EquipmentGanttAssignmentSyncResult> runAsync(
            Path pythonExecutable,
            Path pythonScriptDir,
            Map<String, String> extraUiEnv,
            Request request,
            Consumer<String> logLine) {
        Objects.requireNonNull(request, "request");
        if (!Files.isRegularFile(request.contractPath())) {
            return CompletableFuture.failedFuture(
                    new IOException("contract not found: " + request.contractPath()));
        }
        if (!Files.isRegularFile(request.updatesJsonPath())) {
            return CompletableFuture.failedFuture(
                    new IOException("updates json not found: " + request.updatesJsonPath()));
        }
        Path script = pythonScriptDir.resolve(SCRIPT);
        if (!Files.isRegularFile(script)) {
            return CompletableFuture.failedFuture(
                    new IOException("missing " + SCRIPT + " in " + pythonScriptDir));
        }
        return CompletableFuture.supplyAsync(
                () -> {
                    try {
                        List<String> cmd = new ArrayList<>();
                        cmd.add(pythonExecutable.toString());
                        cmd.add(script.toAbsolutePath().toString());
                        cmd.addAll(buildArgs(request));
                        ProcessBuilder pb = new ProcessBuilder(cmd);
                        pb.directory(pythonScriptDir.toFile());
                        pb.redirectErrorStream(true);
                        PythonProcessRunner.mergeUiEnvIntoProcess(
                                pb, extraUiEnv, pythonScriptDir);
                        Process p = pb.start();
                        StringBuilder merged = new StringBuilder();
                        try (BufferedReader br =
                                new BufferedReader(
                                        new InputStreamReader(
                                                p.getInputStream(), StandardCharsets.UTF_8))) {
                            String line;
                            while ((line = br.readLine()) != null) {
                                if (logLine != null) {
                                    logLine.accept(line);
                                }
                                if (!merged.isEmpty()) {
                                    merged.append('\n');
                                }
                                merged.append(line);
                            }
                        }
                        int code = p.waitFor();
                        EquipmentGanttAssignmentSyncResult parsed =
                                EquipmentGanttAssignmentSyncResult.parseJson(merged.toString());
                        if (code != 0 && !parsed.hasWarnings()) {
                            throw new IOException(
                                    "sync exit " + code + ": " + parsed.formatIssuesForDialog());
                        }
                        return parsed;
                    } catch (Exception e) {
                        if (e instanceof RuntimeException re) {
                            throw re;
                        }
                        throw new RuntimeException(e);
                    }
                });
    }

    private static List<String> buildArgs(Request request) {
        List<String> args = new ArrayList<>();
        args.add("--contract");
        args.add(request.contractPath().toAbsolutePath().toString());
        args.add("--updates-json");
        args.add(request.updatesJsonPath().toAbsolutePath().toString());
        if (request.planXlsxPath() != null && Files.isRegularFile(request.planXlsxPath())) {
            args.add("--plan-xlsx");
            args.add(request.planXlsxPath().toAbsolutePath().toString());
        }
        if (request.expectedTimelineHash() != null
                && !request.expectedTimelineHash().isBlank()) {
            args.add("--expected-timeline-hash");
            args.add(request.expectedTimelineHash().strip());
        }
        if (request.confirmToken() != null && !request.confirmToken().isBlank()) {
            args.add("--confirm-token");
            args.add(request.confirmToken().strip());
        }
        if (request.forceWarnings()) {
            args.add("--force-warnings");
        }
        if (request.dryRun()) {
            args.add("--dry-run");
        }
        return args;
    }

    public static Path defaultScriptDir(Map<String, String> ui) {
        return AppPaths.resolvePythonScriptDir(ui);
    }
}
