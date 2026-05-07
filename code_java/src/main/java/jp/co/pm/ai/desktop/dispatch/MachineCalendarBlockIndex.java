package jp.co.pm.ai.desktop.dispatch;

import java.io.BufferedReader;
import java.io.File;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * Equipment calendar blocks: days where the machine column has at least one occupied slot (planning_core).
 */
public final class MachineCalendarBlockIndex {

    /**
     * Result of {@link #loadOutcome(Path, Path, Path)}: parsed index, optional Python {@code error}, optional
     * {@code diagnostics} JSON string from the export tool.
     */
    public record LoadOutcome(
            MachineCalendarBlockIndex index, String pythonJsonError, String pythonDiagnosticsJson) {}

    private static final ObjectMapper JSON = new ObjectMapper();

    private final Map<String, Set<LocalDate>> blocksByEquipmentKey;

    public MachineCalendarBlockIndex(Map<String, Set<LocalDate>> blocksByEquipmentKey) {
        this.blocksByEquipmentKey = Map.copyOf(blocksByEquipmentKey);
    }

    public boolean isEmpty() {
        return blocksByEquipmentKey.isEmpty();
    }

    public static MachineCalendarBlockIndex empty() {
        return new MachineCalendarBlockIndex(Map.of());
    }

    /**
     * Runs {@code export_machine_calendar_blocks.py} from {@code pythonScriptDir}.
     *
     * @return empty index on failure (drops allowed everywhere calendar unknown).
     */
    public static MachineCalendarBlockIndex load(Path masterWorkbook, Path pythonExe, Path pythonScriptDir)
            throws Exception {
        return loadOutcome(masterWorkbook, pythonExe, pythonScriptDir).index();
    }

    /**
     * Same as {@link #load} but surfaces {@code error} from Python stdout JSON when present (see {@link LoadOutcome}).
     */
    public static LoadOutcome loadOutcome(Path masterWorkbook, Path pythonExe, Path pythonScriptDir)
            throws Exception {
        String payload = runExportScriptRawPayload(masterWorkbook, pythonExe, pythonScriptDir);
        if (payload == null) {
            return new LoadOutcome(empty(), null, null);
        }
        return parseLoadOutcome(payload);
    }

    /**
     * Reads {@link jp.co.pm.ai.desktop.config.AppPaths#resolveMachineCalendarBlocksJsonPath(java.util.Map)} output:
     * same JSON shape as {@code export_machine_calendar_blocks.py}. {@link LoadOutcome#pythonJsonError()} is
     * {@code missing_file} when the file is absent.
     */
    public static LoadOutcome loadOutcomeFromJsonFile(Path jsonFile) throws Exception {
        if (jsonFile == null || !Files.isRegularFile(jsonFile)) {
            return new LoadOutcome(empty(), "missing_file", null);
        }
        String s = Files.readString(jsonFile, StandardCharsets.UTF_8);
        String trimmed = s.trim();
        return parseLoadOutcome(trimmed);
    }

    /**
     * Runs the Python export script and writes pretty-printed JSON to {@code outputJsonFile} (parent dirs created).
     * Overwrites the file on success.
     */
    public static LoadOutcome exportMasterWorkbookToJsonFile(
            Path masterWorkbook, Path pythonExe, Path pythonScriptDir, Path outputJsonFile) throws Exception {
        String payload = runExportScriptRawPayload(masterWorkbook, pythonExe, pythonScriptDir);
        if (payload == null) {
            return new LoadOutcome(empty(), null, null);
        }
        LoadOutcome lo = parseLoadOutcome(payload);
        Path parent = outputJsonFile.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        JsonNode root = JSON.readTree(payload);
        String pretty = JSON.writerWithDefaultPrettyPrinter().writeValueAsString(root);
        Files.writeString(outputJsonFile, pretty, StandardCharsets.UTF_8, StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING, StandardOpenOption.WRITE);
        return lo;
    }

    /**
     * Like {@link #loadOutcome(Path, Path, Path)} but if the primary workbook yields no blocks and no JSON error, retries
     * with {@code summaryWorkbook} when it exists and differs from {@code primaryWorkbook}. The written JSON reflects
     * the last run (fallback overwrites the same output path).
     */
    public static LoadOutcome exportWithSummaryFallbackToJsonFile(
            Path primaryWorkbook,
            Path summaryWorkbook,
            Path pythonExe,
            Path pythonScriptDir,
            Path outputJsonFile)
            throws Exception {
        LoadOutcome lo =
                exportMasterWorkbookToJsonFile(primaryWorkbook, pythonExe, pythonScriptDir, outputJsonFile);
        if (!lo.index().isEmpty() || lo.pythonJsonError() != null) {
            return lo;
        }
        if (summaryWorkbook == null || !Files.isRegularFile(summaryWorkbook)) {
            return lo;
        }
        Path pNorm = primaryWorkbook.toAbsolutePath().normalize();
        Path sNorm = summaryWorkbook.toAbsolutePath().normalize();
        if (pNorm.equals(sNorm)) {
            return lo;
        }
        return exportMasterWorkbookToJsonFile(summaryWorkbook, pythonExe, pythonScriptDir, outputJsonFile);
    }

    /**
     * @return JSON payload string for {@link #parseLoadOutcome(String)}, or {@code null} if prerequisites/process failed.
     */
    private static String runExportScriptRawPayload(Path masterWorkbook, Path pythonExe, Path pythonScriptDir)
            throws Exception {
        if (masterWorkbook == null || !Files.isRegularFile(masterWorkbook)) {
            return null;
        }
        Path script = pythonScriptDir.resolve("export_machine_calendar_blocks.py");
        if (!Files.isRegularFile(script)) {
            return null;
        }
        Path planningCoreInit = pythonScriptDir.resolve("planning_core").resolve("__init__.py");
        if (!Files.isRegularFile(planningCoreInit)) {
            return syntheticToolErrorJson(
                    "planning_core が見つかりません（pm-ai-data のバンドル不整合の可能性）。期待パス: "
                            + planningCoreInit.toAbsolutePath().normalize());
        }
        ProcessBuilder pb =
                new ProcessBuilder(
                        pythonExe.toString(),
                        script.toString(),
                        masterWorkbook.toAbsolutePath().toString());
        pb.directory(pythonScriptDir.toFile());
        pb.environment().put("PM_AI_ALLOW_LEGACY_PYTHON_FOR_TOOLS", "1");
        // Embedded / subprocess Python may omit the script dir from sys.path; planning_core lives beside this script.
        String pyDirNorm = pythonScriptDir.toAbsolutePath().normalize().toString();
        Map<String, String> env = pb.environment();
        String prevPyPath = env.get("PYTHONPATH");
        env.put(
                "PYTHONPATH",
                (prevPyPath == null || prevPyPath.isBlank())
                        ? pyDirNorm
                        : pyDirNorm + File.pathSeparator + prevPyPath);
        pb.redirectErrorStream(true);
        Process p = pb.start();
        String rawStdout;
        try (BufferedReader br =
                new BufferedReader(new InputStreamReader(p.getInputStream(), StandardCharsets.UTF_8))) {
            StringBuilder sb = new StringBuilder();
            String line;
            while ((line = br.readLine()) != null) {
                sb.append(line).append('\n');
            }
            rawStdout = sb.toString();
        }
        boolean finished = p.waitFor(120, TimeUnit.SECONDS);
        if (!finished) {
            p.destroyForcibly();
            return null;
        }
        if (p.exitValue() != 0) {
            return null;
        }
        return pickJsonPayload(rawStdout);
    }

    /** Same JSON shape as {@code export_machine_calendar_blocks.py} when Python is not run. */
    private static String syntheticToolErrorJson(String errorText) {
        try {
            ObjectNode root = JSON.createObjectNode();
            root.put("error", errorText);
            root.set("blocks", JSON.createObjectNode());
            return JSON.writeValueAsString(root);
        } catch (Exception e) {
            return "{\"error\":\"synthetic_tool_error\",\"blocks\":{}}";
        }
    }

    /**
     * Subprocess may merge stderr into stdout; log lines before JSON break parsing. Takes the last line that starts with
     * {@code '{'}, which is the emitted JSON object from export_machine_calendar_blocks.py.
     */
    static String pickJsonPayload(String stdout) {
        if (stdout == null || stdout.isBlank()) {
            return "{}";
        }
        String[] lines = stdout.split("\\R");
        for (int i = lines.length - 1; i >= 0; i--) {
            String t = lines[i].trim();
            if (!t.isEmpty() && t.charAt(0) == '{') {
                return t;
            }
        }
        return stdout.trim();
    }

    static MachineCalendarBlockIndex parseStdoutJson(String stdout) {
        return parseLoadOutcome(stdout).index();
    }

    /**
     * Parses stdout JSON: {@code blocks} object plus optional {@code error} string (tool diagnostic).
     */
    static LoadOutcome parseLoadOutcome(String stdout) {
        try {
            JsonNode root = JSON.readTree(stdout);
            String pythonErr = extractPythonJsonError(root);
            String diag = extractDiagnosticsJson(root);
            JsonNode blocks = root.get("blocks");
            if (!(blocks instanceof ObjectNode on)) {
                return new LoadOutcome(empty(), pythonErr, diag);
            }
            return new LoadOutcome(indexFromBlocksObject(on), pythonErr, diag);
        } catch (Exception e) {
            return new LoadOutcome(empty(), null, null);
        }
    }

    private static String extractDiagnosticsJson(JsonNode root) {
        if (root == null || !root.has("diagnostics")) {
            return null;
        }
        JsonNode d = root.get("diagnostics");
        if (d == null || d.isNull() || d.isMissingNode()) {
            return null;
        }
        try {
            return JSON.writeValueAsString(d);
        } catch (Exception e) {
            return null;
        }
    }

    private static String extractPythonJsonError(JsonNode root) {
        if (root == null || !root.has("error")) {
            return null;
        }
        JsonNode n = root.get("error");
        if (n == null || !n.isTextual()) {
            return null;
        }
        String t = n.asText("").trim();
        return t.isEmpty() ? null : t;
    }

    private static MachineCalendarBlockIndex indexFromBlocksObject(ObjectNode on) {
        Map<String, Set<LocalDate>> acc = new HashMap<>();
        on.fieldNames()
                .forEachRemaining(
                        key -> {
                            Set<LocalDate> ds = new HashSet<>();
                            JsonNode arr = on.get(key);
                            if (arr != null && arr.isArray()) {
                                for (JsonNode n : arr) {
                                    if (n.isTextual()) {
                                        String t = n.asText("");
                                        if (t.length() >= 10) {
                                            LocalDate d = LocalDate.parse(t.substring(0, 10));
                                            ds.add(d);
                                        }
                                    }
                                }
                            }
                            if (!ds.isEmpty()) {
                                acc.put(key.trim(), ds);
                            }
                        });
        return new MachineCalendarBlockIndex(acc);
    }

    /**
     * Whether dropping allocation quantity onto this process/machine/day should be denied (machine occupied /
     * blocked day).
     */
    public boolean isBlockedDay(String processName, String machineName, LocalDate day) {
        if (blocksByEquipmentKey.isEmpty()) {
            return false;
        }
        String p = nz(processName).trim();
        String m = nz(machineName).trim();
        for (Map.Entry<String, Set<LocalDate>> e : blocksByEquipmentKey.entrySet()) {
            if (!e.getValue().contains(day)) {
                continue;
            }
            if (matchesEquipmentKey(e.getKey(), p, m)) {
                return true;
            }
        }
        return false;
    }

    static boolean matchesEquipmentKey(String equipmentKey, String process, String machine) {
        String k = equipmentKey.trim();
        if (k.isEmpty()) {
            return false;
        }
        if (k.equals(machine)) {
            return true;
        }
        if (k.equals(process + "\t" + machine)) {
            return true;
        }
        String compactK = k.replaceAll("\\s+", "");
        String compactPM = (process + machine).replaceAll("\\s+", "");
        if (!compactPM.isEmpty() && compactK.equals(compactPM)) {
            return true;
        }
        if (!machine.isEmpty() && k.contains(machine.trim())) {
            if (process.isEmpty() || k.contains(process.trim())) {
                return true;
            }
        }
        return false;
    }

    private static String nz(String s) {
        return s != null ? s : "";
    }
}
