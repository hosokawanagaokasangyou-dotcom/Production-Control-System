package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.Comparator;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.Optional;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * Path resolution for the desktop UI. <strong>Does not read {@link System#getenv}</strong>; pass keys from
 * the environment-variable tab via {@code ui} (e.g. {@code PM_AI_CODE_PYTHON_DIR}, {@code PM_AI_REPO_ROOT}).
 */
public final class AppPaths {

    public static final String KEY_PM_AI_PYTHON = "PM_AI_PYTHON";
    public static final String KEY_PM_AI_CODE_PYTHON_DIR = "PM_AI_CODE_PYTHON_DIR";
    public static final String KEY_PM_AI_REPO_ROOT = "PM_AI_REPO_ROOT";
    public static final String KEY_PM_AI_WORKSPACE = "PM_AI_WORKSPACE";
    public static final String KEY_PM_AI_TASK_INPUT_SOURCE_DIR = "PM_AI_TASK_INPUT_SOURCE_DIR";

    /** Folder for machining actual-detail Excel exports (PQ plan/02 {@code Folder.Files}). */
    public static final String KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR = "PM_AI_ACTUAL_DETAIL_SOURCE_DIR";

    /**
     * Output directory for the standalone result dispatch table xlsx (Power Query {@code _q} + file name;
     * named range folder path in Excel). Default: {@code resolveRepoRoot(ui)/code}.
     */
    public static final String KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR = "PM_AI_RESULT_DISPATCH_TABLE_DIR";

    /** Gantt compare: directory containing snapshot subfolders (planning_core). */
    public static final String KEY_COMPARE_GANTT_SNAPSHOT_DIR = "COMPARE_GANTT_SNAPSHOT_DIR";

    /**
     * Encrypted Gemini credentials JSON path ({@code gemini_credentials.encrypted.json}); passed to Python
     * {@code GEMINI_CREDENTIALS_JSON}.
     */
    public static final String KEY_GEMINI_CREDENTIALS_JSON = "GEMINI_CREDENTIALS_JSON";

    /**
     * UTF-8 JSON for exclude rules; optional alternative to Excel
     * {@code \u8a2d\u5b9a_\u914d\u53f0\u4e0d\u8981\u5de5\u7a0b}.
     */
    public static final String KEY_PM_AI_EXCLUDE_RULES_JSON = "PM_AI_EXCLUDE_RULES_JSON";

    /** Absolute path to master workbook ({@code master.xlsm}); overrides basename-only {@code MASTER_WORKBOOK_FILE}. */
    public static final String KEY_PM_AI_MASTER_WORKBOOK = "PM_AI_MASTER_WORKBOOK";

    /** Basename or relative master workbook filename (same as {@code MASTER_WORKBOOK_FILE} / planning_core). */
    public static final String KEY_MASTER_WORKBOOK_FILE = "MASTER_WORKBOOK_FILE";

    /**
     * Workbook containing {@code \u5217\u8a2d\u5b9a_\u7d50\u679c_\u30bf\u30b9\u30af\u4e00\u89a7} (optional override from
     * TASK_INPUT_WORKBOOK).
     */
    public static final String KEY_PM_AI_COLUMN_CONFIG_WORKBOOK = "PM_AI_COLUMN_CONFIG_WORKBOOK";

    /** Workbook for plan-sheet data-extraction timestamp columns (optional). */
    public static final String KEY_PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK = "PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK";

    /** CSV for result-task column visibility/order ({@code PM_AI_RESULT_TASK_COLUMN_CONFIG_CSV}). */
    public static final String KEY_PM_AI_RESULT_TASK_COLUMN_CONFIG_CSV = "PM_AI_RESULT_TASK_COLUMN_CONFIG_CSV";

    /**
     * When truthy, {@code workbook_env_bootstrap} skips reading the macro book
     * {@code \u8a2d\u5b9a_\u74b0\u5883\u5909\u6570} sheet (JavaFX tab is source of truth for the child process).
     */
    public static final String KEY_PM_AI_SKIP_WORKBOOK_ENV_SHEET = "PM_AI_SKIP_WORKBOOK_ENV_SHEET";

    /**
     * Env keys whose value is a directory (folder picker in the UI).
     */
    private static final Set<String> FOLDER_PATH_ENV_KEYS = Set.of(
            KEY_PM_AI_CODE_PYTHON_DIR,
            KEY_PM_AI_REPO_ROOT,
            KEY_PM_AI_WORKSPACE,
            KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
            KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR,
            KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR,
            KEY_COMPARE_GANTT_SNAPSHOT_DIR);

    /** Env keys whose value is a single file path (file chooser in the UI). */
    private static final Set<String> FILE_PATH_ENV_KEYS = Set.of(
            KEY_GEMINI_CREDENTIALS_JSON,
            KEY_PM_AI_EXCLUDE_RULES_JSON,
            KEY_PM_AI_MASTER_WORKBOOK,
            KEY_PM_AI_COLUMN_CONFIG_WORKBOOK,
            KEY_PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK,
            KEY_PM_AI_RESULT_TASK_COLUMN_CONFIG_CSV);

    private AppPaths() {}

    // #region agent log
    private static final Path AGENT_DEBUG_LOG =
            Path.of("/mnt/c/\u5de5\u7a0b\u7ba1\u7406AI\u30d7\u30ed\u30b8\u30a7\u30af\u30c8_JAVA/.cursor/debug-e6e382.log");

    private static String agentJsonEsc(String s) {
        if (s == null) {
            return "";
        }
        return s.replace("\\", "\\\\").replace("\"", "\\\"");
    }

    private static void agentLog(String hypothesisId, String location, String message, String dataJson) {
        try {
            long ts = System.currentTimeMillis();
            String line =
                    "{\"sessionId\":\"e6e382\",\"hypothesisId\":\""
                            + hypothesisId
                            + "\",\"location\":\""
                            + agentJsonEsc(location)
                            + "\",\"message\":\""
                            + agentJsonEsc(message)
                            + "\",\"timestamp\":"
                            + ts
                            + ",\"runId\":\"post-fix\",\"data\":"
                            + (dataJson != null ? dataJson : "{}")
                            + "}\n";
            Files.writeString(
                    AGENT_DEBUG_LOG,
                    line,
                    StandardCharsets.UTF_8,
                    StandardOpenOption.CREATE,
                    StandardOpenOption.APPEND);
        } catch (Exception ignored) {
        }
    }

    // #endregion

    /** Whether {@code key} refers to a folder path (not a single file). */
    public static boolean isFolderPathEnvKey(String key) {
        return key != null && FOLDER_PATH_ENV_KEYS.contains(key.trim());
    }

    /** Whether {@code key} refers to a file path (encrypted JSON etc.). */
    public static boolean isFilePathEnvKey(String key) {
        return key != null && FILE_PATH_ENV_KEYS.contains(key.trim());
    }

    /** JSON credentials or exclude-rules file ({@code *.json}). */
    public static boolean isJsonFilePathEnvKey(String key) {
        String k = key != null ? key.trim() : "";
        return KEY_GEMINI_CREDENTIALS_JSON.equals(k) || KEY_PM_AI_EXCLUDE_RULES_JSON.equals(k);
    }

    /** Master / column-config / data-extraction workbooks ({@code *.xlsm}, {@code *.xlsx}). */
    public static boolean isExcelWorkbookPathEnvKey(String key) {
        String k = key != null ? key.trim() : "";
        return KEY_PM_AI_MASTER_WORKBOOK.equals(k)
                || KEY_PM_AI_COLUMN_CONFIG_WORKBOOK.equals(k)
                || KEY_PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK.equals(k);
    }

    /** Result-task column config CSV. */
    public static boolean isCsvFilePathEnvKey(String key) {
        return key != null && KEY_PM_AI_RESULT_TASK_COLUMN_CONFIG_CSV.equals(key.trim());
    }

    /**
     * Same UNC as {@code plan/01_*.m} {@code \u30d1\u30b9} (PQ-A / production plan inquiry folder).
     */
    private static final String DEFAULT_PQ_A_SOURCE_UNC =
            "\\\\192.168.0.101\\"
                    + "\u5171\u6709\u30d5\u30a9\u30eb\u30c0\\"
                    + "\u6e56\u5357\u5de5\u5834\\"
                    + "\u6e56\u5357\u5171\u6709\\"
                    + "\u751f\u7523\u7ba1\u7406\u30b7\u30b9\u30c6\u30e0\\"
                    + "\u7ba1\u7406\u30b7\u30b9\u30c6\u30e0\\"
                    + "\u25cfDATA\\"
                    + "\u751f\u7523\u8a08\u753b\u554f\u5408\u305b";

    /** Same as {@code plan/02__q\u52a0\u5de5\u5b9f\u7e3e\u660e\u7d30DATA.m} {@code Folder.Files} path. */
    private static final String DEFAULT_ACTUAL_DETAIL_SOURCE_UNC =
            "\\\\192.168.0.101\\"
                    + "\u5171\u6709\u30d5\u30a9\u30eb\u30c0\\"
                    + "\u6e56\u5357\u5de5\u5834\\"
                    + "\u6e56\u5357\u5171\u6709\\"
                    + "002  \u52a0\u5de5G\\"
                    + "\u25cf\u691c\u67fb\u8868\u4f5c\u6210\\"
                    + "\u52a0\u5de5\u5b9f\u7e3e\u660e\u7d30DATA";

    /**
     * {@code ui} from the env tab; {@code null} or empty map uses directory walk only (no overrides).
     */
    public static Path resolvePythonScriptDir(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String override = trim(u.get(KEY_PM_AI_CODE_PYTHON_DIR));
        if (!override.isEmpty()) {
            Path p = Path.of(override).toAbsolutePath().normalize();
            if (Files.isDirectory(p)) {
                return p;
            }
        }
        String repo = trim(u.get(KEY_PM_AI_REPO_ROOT));
        if (!repo.isEmpty()) {
            Path base = Path.of(repo).toAbsolutePath().normalize();
            Path underRepo = base.resolve("code").resolve("python");
            Path underNested = base.resolve("Production-Control-System").resolve("code").resolve("python");
            for (Path p : new Path[] {underRepo, underNested}) {
                if (Files.isDirectory(p) && Files.isRegularFile(p.resolve("task_extract_stage1.py"))) {
                    agentLog(
                            "H1",
                            "AppPaths.resolvePythonScriptDir",
                            "chosen with task_extract_stage1.py",
                            "{\"chosen\":\""
                                    + agentJsonEsc(p.toString())
                                    + "\",\"underRepo\":\""
                                    + agentJsonEsc(underRepo.toString())
                                    + "\",\"underNested\":\""
                                    + agentJsonEsc(underNested.toString())
                                    + "\"}");
                    return p;
                }
            }
            for (Path p : new Path[] {underRepo, underNested}) {
                if (Files.isDirectory(p)) {
                    agentLog(
                            "H1",
                            "AppPaths.resolvePythonScriptDir",
                            "chosen directory only",
                            "{\"chosen\":\""
                                    + agentJsonEsc(p.toString())
                                    + "\",\"underRepo\":\""
                                    + agentJsonEsc(underRepo.toString())
                                    + "\",\"underNested\":\""
                                    + agentJsonEsc(underNested.toString())
                                    + "\"}");
                    return p;
                }
            }
        }
        Path start = Path.of(System.getProperty("user.dir", ".")).toAbsolutePath().normalize();
        Optional<Path> found = findCodePythonFrom(start);
        if (found.isPresent()) {
            return found.get();
        }
        Path sibling = start.resolve("..").resolve("code").resolve("python").normalize();
        if (Files.isDirectory(sibling)) {
            return sibling;
        }
        return sibling;
    }

    /** PQ-A task-input folder; optional {@code PM_AI_TASK_INPUT_SOURCE_DIR} in {@code ui}. */
    public static Path resolveTaskInputSourceDir(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String override = trim(u.get(KEY_PM_AI_TASK_INPUT_SOURCE_DIR));
        if (!override.isEmpty()) {
            return Path.of(override).toAbsolutePath().normalize();
        }
        return Path.of(DEFAULT_PQ_A_SOURCE_UNC);
    }

    /** Machining actual-detail export folder; optional {@code PM_AI_ACTUAL_DETAIL_SOURCE_DIR} in {@code ui}. */
    public static Path resolveActualDetailSourceDir(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String override = trim(u.get(KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR));
        if (!override.isEmpty()) {
            return Path.of(override).toAbsolutePath().normalize();
        }
        return Path.of(DEFAULT_ACTUAL_DETAIL_SOURCE_UNC);
    }

    /**
     * Directory for standalone result-dispatch xlsx; optional {@code PM_AI_RESULT_DISPATCH_TABLE_DIR} in
     * {@code ui}. Default: {@code resolveRepoRoot(ui)}/{@code code} (e.g. Production-Control-System/code).
     */
    public static Path resolveResultDispatchTableDir(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String override = trim(u.get(KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR));
        if (!override.isEmpty()) {
            return Path.of(override).toAbsolutePath().normalize();
        }
        return resolveRepoRoot(u).resolve("code").toAbsolutePath().normalize();
    }

    /**
     * First existing {@code master.xlsm} / {@code master.xlsx} under {@link #resolveRepoRoot(Map)} ({@code plan/},
     * {@code code/}, or repo root). Used for JavaFX bootstrap hints only.
     */
    public static Optional<Path> resolveMasterWorkbookCandidate(Map<String, String> ui) {
        Path root = resolveRepoRoot(ui != null ? ui : Map.of());
        Path[] candidates =
                new Path[] {
                    root.resolve("plan").resolve("master.xlsm"),
                    root.resolve("plan").resolve("master.xlsx"),
                    root.resolve("code").resolve("master.xlsm"),
                    root.resolve("master.xlsm"),
                };
        for (Path c : candidates) {
            if (Files.isRegularFile(c)) {
                return Optional.of(c.toAbsolutePath().normalize());
            }
        }
        return Optional.empty();
    }

    /**
     * Approximates {@code planning_core} bootstrap {@code os.getcwd()} after import: {@code PM_AI_WORKSPACE}
     * if set and a directory, else parent of {@code TASK_INPUT_WORKBOOK}, else parent of {@link
     * #resolvePythonScriptDir(Map)} (the {@code code} folder next to {@code python}).
     */
    public static Path resolveEffectivePlanningCwd(Map<String, String> ui, String taskInputWorkbookPath) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String ws = trim(u.get(KEY_PM_AI_WORKSPACE));
        if (!ws.isEmpty()) {
            Path w = Path.of(ws).toAbsolutePath().normalize();
            if (Files.isDirectory(w)) {
                return w;
            }
        }
        String tb = taskInputWorkbookPath != null ? taskInputWorkbookPath.trim() : "";
        if (!tb.isEmpty()) {
            Path p = Path.of(tb).toAbsolutePath().normalize();
            Path parent = p.getParent();
            if (parent != null && Files.isDirectory(parent)) {
                return parent;
            }
        }
        Path py = resolvePythonScriptDir(u);
        Path codeDir = py.getParent();
        if (codeDir != null && Files.isDirectory(codeDir)) {
            return codeDir.toAbsolutePath().normalize();
        }
        return resolveRepoRoot(u).toAbsolutePath().normalize();
    }

    /**
     * Same resolution as {@code planning_core._core._master_workbook_path_resolved} for the given env and
     * effective macro-book path.
     */
    public static Path resolveMasterWorkbookPathResolved(Map<String, String> ui, String taskInputWorkbookPath) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String alt = trim(u.get(KEY_PM_AI_MASTER_WORKBOOK));
        if (!alt.isEmpty()) {
            Path ap = Path.of(alt).toAbsolutePath().normalize();
            if (Files.isRegularFile(ap)) {
                return ap;
            }
        }
        String mf = trim(u.get(KEY_MASTER_WORKBOOK_FILE));
        if (mf.isEmpty()) {
            mf = "master.xlsm";
        }
        Path cwd = resolveEffectivePlanningCwd(u, taskInputWorkbookPath);
        if (mf.startsWith("\\\\")) {
            return Path.of(mf);
        }
        Path mfPath = Path.of(mf);
        if (mfPath.isAbsolute()) {
            return mfPath.normalize();
        }
        return cwd.resolve(mf).normalize().toAbsolutePath();
    }

    /** Filename for stage-1 shaped tasks ({@code planning_core.STAGE1_OUTPUT_FILENAME}). */
    public static final String STAGE1_PLAN_TASKS_FILENAME = "plan_input_tasks.xlsx";

    /** Sheet name in {@link #STAGE1_PLAN_TASKS_FILENAME} ({@code planning_core.run_stage1_extract} / {@code to_excel}). */
    public static final String STAGE1_PLAN_OUTPUT_SHEET = "\u30bf\u30b9\u30af\u4e00\u89a7";

    /**
     * Preview workbook written right after {@code load_tasks_df} ({@code planning_core.STAGE1_TASK_INPUT_PREVIEW_FILENAME}).
     */
    public static final String STAGE1_TASK_INPUT_PREVIEW_FILENAME = "stage1_task_input_table.xlsx";

    /** Sheet name inside {@link #STAGE1_TASK_INPUT_PREVIEW_FILENAME}. */
    public static final String STAGE1_TASK_INPUT_PREVIEW_SHEET = "\u30bf\u30b9\u30af\u5165\u529b\u6574\u5f62";

    /**
     * Written by {@code run_stage1_extract} beside {@code json_data_dir} ({@code planning_core} /
     * {@code STAGE1_EXCLUDE_RULES_JSON_FILENAME}).
     */
    public static final String STAGE1_EXCLUDE_RULES_JSON_FILENAME = "stage1_exclude_rules.json";

    /**
     * Path to the stage-1 exclude-rules sidecar JSON (same convention as Python {@code json_data_dir} under
     * {@code code/python}).
     */
    public static Path stage1ExcludeRulesJsonPath(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        return resolvePythonScriptDir(u)
                .resolve("json")
                .resolve(STAGE1_EXCLUDE_RULES_JSON_FILENAME)
                .toAbsolutePath()
                .normalize();
    }

    /**
     * Default path to stage-1 Excel output.
     *
     * <p>{@code planning_core.bootstrap} sets {@code cwd} to the parent of {@code code/python} (or
     * {@code PM_AI_WORKSPACE}), so Python writes to {@code <that cwd>/output/plan_input_tasks.xlsx}
     * i.e. typically {@code Production-Control-System/code/output/}, not {@code code/python/output/}.
     */
    public static Path defaultStage1PlanTasksPath(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path pyDir = resolvePythonScriptDir(u);
        Path parent = pyDir.getParent();
        Path underCodeOutput =
                parent != null
                        ? parent.resolve("output").resolve(STAGE1_PLAN_TASKS_FILENAME)
                        : pyDir.resolve("output").resolve(STAGE1_PLAN_TASKS_FILENAME);
        Path underPyOutput = pyDir.resolve("output").resolve(STAGE1_PLAN_TASKS_FILENAME);
        if (Files.isRegularFile(underCodeOutput)) {
            return underCodeOutput.toAbsolutePath().normalize();
        }
        if (Files.isRegularFile(underPyOutput)) {
            return underPyOutput.toAbsolutePath().normalize();
        }
        Path repo = resolveRepoRoot(u);
        Path underCodePython =
                repo.resolve("code").resolve("python").resolve("output").resolve(STAGE1_PLAN_TASKS_FILENAME);
        if (Files.isRegularFile(underCodePython)) {
            return underCodePython.toAbsolutePath().normalize();
        }
        String ws = trim(u.get(KEY_PM_AI_WORKSPACE));
        if (!ws.isEmpty()) {
            Path w = Path.of(ws).resolve("output").resolve(STAGE1_PLAN_TASKS_FILENAME);
            if (Files.isRegularFile(w)) {
                return w.toAbsolutePath().normalize();
            }
        }
        return underCodeOutput.toAbsolutePath().normalize();
    }

    /**
     * Default path to the stage-1 task-input preview xlsx (tabular state after header cleanup, before plan_input_tasks).
     */
    public static Path defaultStage1TaskInputPreviewPath(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path pyDir = resolvePythonScriptDir(u);
        Path parent = pyDir.getParent();
        Path underCodeOutput =
                parent != null
                        ? parent.resolve("output").resolve(STAGE1_TASK_INPUT_PREVIEW_FILENAME)
                        : pyDir.resolve("output").resolve(STAGE1_TASK_INPUT_PREVIEW_FILENAME);
        Path underPyOutput = pyDir.resolve("output").resolve(STAGE1_TASK_INPUT_PREVIEW_FILENAME);
        if (Files.isRegularFile(underCodeOutput)) {
            return underCodeOutput.toAbsolutePath().normalize();
        }
        if (Files.isRegularFile(underPyOutput)) {
            return underPyOutput.toAbsolutePath().normalize();
        }
        Path repo = resolveRepoRoot(u);
        Path underCodePython =
                repo.resolve("code").resolve("python").resolve("output").resolve(STAGE1_TASK_INPUT_PREVIEW_FILENAME);
        if (Files.isRegularFile(underCodePython)) {
            return underCodePython.toAbsolutePath().normalize();
        }
        String ws = trim(u.get(KEY_PM_AI_WORKSPACE));
        if (!ws.isEmpty()) {
            Path w = Path.of(ws).resolve("output").resolve(STAGE1_TASK_INPUT_PREVIEW_FILENAME);
            if (Files.isRegularFile(w)) {
                return w.toAbsolutePath().normalize();
            }
        }
        return underCodeOutput.toAbsolutePath().normalize();
    }

    /** Repository root containing {@code code/python}. */
    public static Path resolveRepoRoot(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String r = trim(u.get(KEY_PM_AI_REPO_ROOT));
        if (!r.isEmpty()) {
            return Path.of(r).toAbsolutePath().normalize();
        }
        Path py = resolvePythonScriptDir(u);
        Path code = py.getParent();
        if (code == null) {
            return py;
        }
        Path repo = code.getParent();
        return repo != null ? repo : code;
    }

    /**
     * Discovers a macro {@code .xlsm} for auto-fill (JavaFX main field). Does not read {@code TASK_INPUT_WORKBOOK}
     * from {@code ui} ? that env var is set only by the {@link jp.co.pm.ai.desktop.bridge.PythonProcessRunner}
     * from the main workbook field. Uses {@code PM_AI_WORKSPACE} then {@link #resolveRepoRoot(Map)} scan.
     */
    public static Optional<Path> resolveTaskInputWorkbook(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String ws = trim(u.get(KEY_PM_AI_WORKSPACE));
        if (!ws.isEmpty()) {
            Path w = Path.of(ws).toAbsolutePath().normalize();
            Optional<Path> fromWs = pickMacroWorkbook(w);
            if (fromWs.isPresent()) {
                return fromWs;
            }
        }
        return pickMacroWorkbook(resolveRepoRoot(u));
    }

    private static String trim(String s) {
        return s != null ? s.trim() : "";
    }

    /**
     * Lists {@code .xlsm} in a directory; if one file, returns it; if several, prefers a name
     * containing {@code \u914d\u53f0}, else lexicographically first.
     */
    static Optional<Path> pickMacroWorkbook(Path directory) {
        if (directory == null || !Files.isDirectory(directory)) {
            return Optional.empty();
        }
        final java.util.List<Path> xlsms;
        try (Stream<Path> stream = Files.list(directory)) {
            xlsms = stream
                    .filter(p -> Files.isRegularFile(p)
                            && p.getFileName()
                                    .toString()
                                    .toLowerCase(Locale.ROOT)
                                    .endsWith(".xlsm"))
                    .collect(Collectors.toList());
        } catch (IOException e) {
            return Optional.empty();
        }
        if (xlsms.isEmpty()) {
            return Optional.empty();
        }
        if (xlsms.size() == 1) {
            return Optional.of(xlsms.get(0));
        }
        String marker = "\u914d\u53f0";
        Optional<Path> preferred = xlsms.stream()
                .filter(p -> p.getFileName().toString().contains(marker))
                .min(Comparator.comparing(p -> p.getFileName().toString()));
        return preferred.or(() -> xlsms.stream()
                .min(Comparator.comparing(p -> p.getFileName().toString())));
    }

    private static Optional<Path> findCodePythonFrom(Path start) {
        Path cur = start;
        for (int i = 0; i < 8; i++) {
            Path candidate = cur.resolve("code").resolve("python");
            if (Files.isDirectory(candidate) && Files.isRegularFile(candidate.resolve("task_extract_stage1.py"))) {
                return Optional.of(candidate.toAbsolutePath().normalize());
            }
            Path parent = cur.getParent();
            if (parent == null || Objects.equals(parent, cur)) {
                break;
            }
            cur = parent;
        }
        return Optional.empty();
    }
}
