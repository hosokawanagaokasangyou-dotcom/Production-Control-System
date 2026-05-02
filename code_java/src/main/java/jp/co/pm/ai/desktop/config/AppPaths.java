package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
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

    /** UTF-8 JSON for exclude rules (?i?K1??u???_?z??s?v?H???v????); optional alternative to Excel sheet. */
    public static final String KEY_PM_AI_EXCLUDE_RULES_JSON = "PM_AI_EXCLUDE_RULES_JSON";

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
            KEY_GEMINI_CREDENTIALS_JSON, KEY_PM_AI_EXCLUDE_RULES_JSON);

    private AppPaths() {}

    /** Whether {@code key} refers to a folder path (not a single file). */
    public static boolean isFolderPathEnvKey(String key) {
        return key != null && FOLDER_PATH_ENV_KEYS.contains(key.trim());
    }

    /** Whether {@code key} refers to a file path (encrypted JSON etc.). */
    public static boolean isFilePathEnvKey(String key) {
        return key != null && FILE_PATH_ENV_KEYS.contains(key.trim());
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
            Path p = Path.of(repo, "Production-Control-System", "code", "python");
            if (Files.isDirectory(p)) {
                return p.toAbsolutePath().normalize();
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
