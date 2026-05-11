package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.InvalidPathException;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.LinkedHashMap;
import java.util.Optional;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * Path resolution for the desktop UI. <strong>Does not read {@link System#getenv}</strong>; pass keys from
 * the environment-variable tab via {@code ui} (e.g. {@code PM_AI_CODE_PYTHON_DIR}, {@code PM_AI_REPO_ROOT},
 * {@link #KEY_PM_AI_OUTPUT_DIR}).
 */
public final class AppPaths {

    public static final String KEY_PM_AI_PYTHON = "PM_AI_PYTHON";
    public static final String KEY_PM_AI_CODE_PYTHON_DIR = "PM_AI_CODE_PYTHON_DIR";
    public static final String KEY_PM_AI_REPO_ROOT = "PM_AI_REPO_ROOT";
    public static final String KEY_PM_AI_WORKSPACE = "PM_AI_WORKSPACE";

    /**
     * Cursor デバッグ用 NDJSON ログファイルへの絶対パス（任意）。未設定時は {@link jp.co.pm.ai.desktop.debug.AgentDebugLog}
     * が {@code リポジトリ親/.cursor/debug-&lt;session&gt;.log} などを試す。
     */
    public static final String KEY_PM_AI_CURSOR_DEBUG_LOG = "PM_AI_CURSOR_DEBUG_LOG";

    /**
     * NDJSON デバッグログの追加ミラー先（任意）。Windows JVM が {@code C:\...} に書いた行を、UNC（{@code \\wsl$\...}）など
     * Cursor（WSL）側が読むパスへ複製する場合に使用。{@link jp.co.pm.ai.desktop.debug.AgentDebugLog} を参照。
     */
    public static final String KEY_PM_AI_DEBUG_LOG_MIRROR = "PM_AI_DEBUG_LOG_MIRROR";

    /**
     * Stage1/2 成果物フォルダ（従来の {@code code/output} に相当）。未設定時は {@link #resolveRepoRoot(Map)} の直下の
     * {@code output/}。Python {@code planning_core.bootstrap} の {@code output_dir} と揃える。
     */
    public static final String KEY_PM_AI_OUTPUT_DIR = "PM_AI_OUTPUT_DIR";

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
     * {@code 設定_配台不要工程}.
     */
    public static final String KEY_PM_AI_EXCLUDE_RULES_JSON = "PM_AI_EXCLUDE_RULES_JSON";

    /** リポジトリ直下の人間向け要約（デスクトップ「特別ルール」タブと運用で同期）。 */
    public static final String SPECIAL_RULES_SUMMARY_MD = "特別ルール.md";

    /** リポジトリ直下の L 番号列挙（{@code planning_core/_core.py} のコメントと対応）。 */
    public static final String SPECIAL_RULES_ENUMERATED_MD = "特別ルール列挙.md";

    /** Absolute path to master workbook ({@code master.xlsm}); overrides basename-only {@code MASTER_WORKBOOK_FILE}. */
    public static final String KEY_PM_AI_MASTER_WORKBOOK = "PM_AI_MASTER_WORKBOOK";

    /** Basename or relative master workbook filename (same as {@code MASTER_WORKBOOK_FILE} / planning_core). */
    public static final String KEY_MASTER_WORKBOOK_FILE = "MASTER_WORKBOOK_FILE";

    /**
     * {@code 実行・ログ} タブの「開く」が開くサマリ用マクロブック（
     * 絶対パス、または {@code code/} からの相対）。空で
     * {@link #SUMMARY_AI_DISPATCH_XLSM}。
     */
    public static final String KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK =
            "PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK";

    /**
     * Workbook containing {@code 列設定_結果_タスク一覧} (optional override when
     * it differs from {@code PM_AI_PLAN_INPUT_PATH}).
     */
    public static final String KEY_PM_AI_COLUMN_CONFIG_WORKBOOK = "PM_AI_COLUMN_CONFIG_WORKBOOK";

    /** Workbook for plan-sheet data-extraction timestamp columns (optional). */
    public static final String KEY_PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK = "PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK";

    /** CSV for result-task column visibility/order ({@code PM_AI_RESULT_TASK_COLUMN_CONFIG_CSV}). */
    public static final String KEY_PM_AI_RESULT_TASK_COLUMN_CONFIG_CSV = "PM_AI_RESULT_TASK_COLUMN_CONFIG_CSV";

    /**
     * Plan-input workbook path ({@code PM_AI_PLAN_INPUT_PATH}); CSV / Parquet / Excel. Align with
     * {@link jp.co.pm.ai.desktop.PlanInputTabController}.
     */
    public static final String KEY_PM_AI_PLAN_INPUT_PATH = "PM_AI_PLAN_INPUT_PATH";

    /**
     * Stage1 加工計画DATA相当の単一ファイル（{@code PM_AI_PROCESSING_PLAN_PATH}）。未設定時は Python が
     * {@link #KEY_PM_AI_TASK_INPUT_SOURCE_DIR} 内の最新表を選択する。
     */
    public static final String KEY_PM_AI_PROCESSING_PLAN_PATH = "PM_AI_PROCESSING_PLAN_PATH";

    /**
     * Single-file override for actual-detail workbook ({@code PM_AI_ACTUAL_DETAIL_WORKBOOK}); takes precedence over
     * {@link #KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR} when set.
     */
    public static final String KEY_PM_AI_ACTUAL_DETAIL_WORKBOOK = "PM_AI_ACTUAL_DETAIL_WORKBOOK";

    /** Optional sheet name inside {@link #KEY_PM_AI_ACTUAL_DETAIL_WORKBOOK} (empty = first sheet). */
    public static final String KEY_PM_AI_ACTUAL_DETAIL_SHEET = "PM_AI_ACTUAL_DETAIL_SHEET";

    /**
     * 加工実績明細の元ファイル（Excel/CSV）を読む前のサイズ上限（バイト）。超過時は読込を中止してヒープ枯渇を防ぐ。
     * 空または未設定で {@link #DEFAULT_PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES}。0 以下で上限なし（チェックしない）。
     */
    public static final String KEY_PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES = "PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES";

    /** {@link #KEY_PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES} の既定（20 MiB）。 */
    public static final long DEFAULT_PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES = 20L * 1024 * 1024;

    /** Optional absolute path to result-task JSON sidecar ({@code PM_AI_PLAN_RESULT_TASK_JSON_PATH}). */
    public static final String KEY_PM_AI_PLAN_RESULT_TASK_JSON_PATH = "PM_AI_PLAN_RESULT_TASK_JSON_PATH";

    /**
     * 結果_タスク一覧のサイドカー JSON（{@code PM_AI_PLAN_RESULT_TASK_JSON_PATH} と対）。{@code 0} / {@code false} /
     * {@code no} / {@code off} / {@code none} で無効。未設定で有効（plan_workbook_sidecar）。
     */
    public static final String KEY_PM_AI_PLAN_RESULT_TASK_JSON = "PM_AI_PLAN_RESULT_TASK_JSON";

    /**
     * planning_core master-data table paths ({@code ui_ref_env_defaults.json}): each names a file (CSV / text), not a
     * directory.
     */
    private static final Set<String> TABULAR_DATA_TABLE_PATH_KEYS =
            Set.of(
                    "RAW_FABRIC_WIDTH_TABLE_PATH",
                    "ROLL_UNIT_BY_USED_RAW_TABLE_PATH",
                    "PRODUCT_WIDTH_TABLE_PATH",
                    "PRODUCT_LENGTH_TABLE_PATH",
                    "PRODUCT_THICKNESS_TABLE_PATH");

    /**
     * When truthy, {@code workbook_env_bootstrap} skips reading the macro book
     * {@code 設定_環境変数} sheet (JavaFX tab is source of truth for the child process).
     */
    public static final String KEY_PM_AI_SKIP_WORKBOOK_ENV_SHEET = "PM_AI_SKIP_WORKBOOK_ENV_SHEET";

    /**
     * Python planning_core: mirror {@code 計画*.xlsx} to same-name {@code .json}. Values
     * {@code 0}/{@code false}/{@code no}/{@code off}/{@code none} disable; unset defaults to enabled.
     */
    public static final String KEY_PM_AI_PLAN_WORKBOOK_JSON = "PM_AI_PLAN_WORKBOOK_JSON";

    /**
     * Python planning_core: mirror {@code 人員*.xlsx} to same-name {@code .json}. Same disable tokens as
     * {@link #KEY_PM_AI_PLAN_WORKBOOK_JSON}; unset defaults to enabled.
     */
    public static final String KEY_PM_AI_MEMBER_SCHEDULE_JSON = "PM_AI_MEMBER_SCHEDULE_JSON";

    /**
     * 段階2: {@code 計画*.xlsx} / {@code 人員*.xlsx} を成果物として残す。
     * {@code 0} / {@code false} / {@code no} / {@code off} / {@code none} のときは JSON のみ（UI 実行・ログタブのチェックボックスから上書き可）。
     */
    public static final String KEY_PM_AI_STAGE2_WRITE_EXCEL = "PM_AI_STAGE2_WRITE_EXCEL";

    /**
     * 段階2の Excel 成果物（結果ブック）のフォントファミリ。空のときは planning_core の {@code RESULT_BOOK_FONT_NAME}（BIZ
     * UDゴシック）相当。JavaFX 実行タブのコンボで上書き可。
     */
    public static final String KEY_PM_AI_RESULT_BOOK_FONT = "PM_AI_RESULT_BOOK_FONT";

    /**
     * 段階2の Excel 生成デバッグ: 1 件の依頼NO（例 {@code Y5-14}）を追跡し NDJSON を planning_core から出力。JavaFX
     * {@code 環境変数} タブに設定。空で無効。
     */
    public static final String KEY_PM_AI_EXCEL_TRACE_TASK_ID = "PM_AI_EXCEL_TRACE_TASK_ID";

    /**
     * Windows CLI のエラー後 ``pause``／Enter 待ち（{@code workbook_env_bootstrap.pause_cmd_window_on_cli_error}）。JavaFX
     * からパイプ接続で起動する子プロセスでは stdin が TTY でないため {@code pause} がブロックし得る。未設定時はシェル側で
     * {@code 0} を付与して無効化する。
     */
    public static final String KEY_PM_AI_CMD_PAUSE_ON_ERROR = "PM_AI_CMD_PAUSE_ON_ERROR";

    /**
     * ポータブル配布（{@code pm-ai-data}）の正本。推奨はバージョンアップ用 ZIP、手入力では展開済み正本フォルダも使える。
     * {@link #VERSION_TXT_FILE_NAME} で版比較し、新しいときのみ起動時同期する。
     */
    public static final String KEY_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR = "PM_AI_PORTABLE_BUNDLE_SOURCE_DIR";

    /**
     * 環境変数タブ・{@code ui_ref_env_defaults.json} で値が空のときの工場共有上の正本（UNC）。バージョンアップ用 ZIP と外付け
     * {@code version.txt} を置く {@code pm-ai-package-release} フォルダを指す。ユーザーが上書き可能。
     */
    public static final String DEFAULT_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR =
            "\\\\192.168.0.101\\共有フォルダ\\湖南工場\\湖南共有\\002  加工G\\●配台AIシステム\\pm-ai-package-release\\PMD_version_upgrade.zip";

    /**
     * {@link #KEY_PM_AI_TASK_INPUT_SOURCE_DIR} が空のときの既定（工場共有・生産計画問合せフォルダ）。{@code plan/01_*.m} のパスと揃える。
     */
    public static final String DEFAULT_PM_AI_TASK_INPUT_SOURCE_DIR =
            "\\\\192.168.0.101\\"
                    + "共有フォルダ\\"
                    + "湖南工場\\"
                    + "湖南共有\\"
                    + "生産管理システム\\"
                    + "管理システム\\"
                    + "●DATA\\"
                    + "生産計画問合せ";

    /**
     * {@link #KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR} が空のときの既定（加工実績明細DATA）。{@code plan/02__q加工実績明細DATA.m} と揃える。
     */
    public static final String DEFAULT_PM_AI_ACTUAL_DETAIL_SOURCE_DIR =
            "\\\\192.168.0.101\\"
                    + "共有フォルダ\\"
                    + "湖南工場\\"
                    + "湖南共有\\"
                    + "002  加工G\\"
                    + "●検査表作成\\"
                    + "加工実績明細DATA";

    /** リポジトリ直下および {@code pm-ai-data} 直下で共用する版ファイル名。 */
    public static final String VERSION_TXT_FILE_NAME = "version.txt";

    /**
     * {@code user.dir} 等から同梱パス（{@code pm-ai-data/runtime/python-embed}・{@code code/python}）を探すときの、親ディレクトリ方向の最大ステップ数。
     */
    private static final int BUNDLED_ANCHOR_WALK_MAX_PARENT_HOPS = 12;

    /**
     * 初回インストール用バンドルに同梱する空マーカー（{@code PMD.exe} と同階層）。存在時のみ起動時に環境タブを既定へリセットし、成功後に削除する。
     */
    public static final String PORTABLE_FIRST_LAUNCH_MARKER_FILE = "初回起動.txt";

    /**
     * Env keys whose value is a directory (folder picker in the UI).
     */
    private static final Set<String> FOLDER_PATH_ENV_KEYS = Set.of(
            KEY_PM_AI_CODE_PYTHON_DIR,
            KEY_PM_AI_REPO_ROOT,
            KEY_PM_AI_WORKSPACE,
            KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
            KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR,
            KEY_PM_AI_OUTPUT_DIR,
            KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR,
            KEY_COMPARE_GANTT_SNAPSHOT_DIR);

    /**
     * {@link #normalizedFolderEnvOverrides(Map)} の処理順（{@link #KEY_PM_AI_REPO_ROOT} を先に確定）。
     *
     * <p>{@link #KEY_PM_AI_TASK_INPUT_SOURCE_DIR} / {@link #KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR} はネットワークソース正本のため
     * 含めない（起動後は {@code MainShellController} 側で既定 UNC に固定する）。
     */
    private static final List<String> FOLDER_PATH_NORMALIZE_ORDER =
            List.of(
                    KEY_PM_AI_REPO_ROOT,
                    KEY_PM_AI_CODE_PYTHON_DIR,
                    KEY_PM_AI_WORKSPACE,
                    KEY_PM_AI_OUTPUT_DIR,
                    KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR,
                    KEY_COMPARE_GANTT_SNAPSHOT_DIR);

    /** Env keys whose value is a single file path (file chooser in the UI). */
    private static final Set<String> FILE_PATH_ENV_KEYS = createFilePathEnvKeys();

    private static Set<String> createFilePathEnvKeys() {
        HashSet<String> s = new HashSet<>();
        s.add(KEY_GEMINI_CREDENTIALS_JSON);
        s.add(KEY_PM_AI_EXCLUDE_RULES_JSON);
        s.add(KEY_PM_AI_MASTER_WORKBOOK);
        s.add(KEY_PM_AI_COLUMN_CONFIG_WORKBOOK);
        s.add(KEY_PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK);
        s.add(KEY_PM_AI_RESULT_TASK_COLUMN_CONFIG_CSV);
        s.add(KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK);
        s.add(KEY_PM_AI_PLAN_INPUT_PATH);
        s.add(KEY_PM_AI_PROCESSING_PLAN_PATH);
        s.add(KEY_PM_AI_ACTUAL_DETAIL_WORKBOOK);
        s.add(KEY_PM_AI_PLAN_RESULT_TASK_JSON_PATH);
        s.add(KEY_PM_AI_CURSOR_DEBUG_LOG);
        s.add(KEY_PM_AI_DEBUG_LOG_MIRROR);
        s.add(KEY_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR);
        s.addAll(TABULAR_DATA_TABLE_PATH_KEYS);
        return Set.copyOf(s);
    }

    private AppPaths() {}

    /** Whether {@code key} refers to a folder path (not a single file). */
    public static boolean isFolderPathEnvKey(String key) {
        if (key == null) {
            return false;
        }
        String k = key.trim();
        if (FILE_PATH_ENV_KEYS.contains(k)) {
            return false;
        }
        return FOLDER_PATH_ENV_KEYS.contains(k);
    }

    /** Whether {@code key} refers to a file path (encrypted JSON etc.). */
    public static boolean isFilePathEnvKey(String key) {
        return key != null && FILE_PATH_ENV_KEYS.contains(key.trim());
    }

    /** JSON credentials or exclude-rules file ({@code *.json}). */
    public static boolean isJsonFilePathEnvKey(String key) {
        String k = key != null ? key.trim() : "";
        return KEY_GEMINI_CREDENTIALS_JSON.equals(k)
                || KEY_PM_AI_EXCLUDE_RULES_JSON.equals(k)
                || KEY_PM_AI_PLAN_RESULT_TASK_JSON_PATH.equals(k);
    }

    /** Master / column-config / data-extraction workbooks ({@code *.xlsm}, {@code *.xlsx}). */
    public static boolean isExcelWorkbookPathEnvKey(String key) {
        String k = key != null ? key.trim() : "";
        return KEY_PM_AI_MASTER_WORKBOOK.equals(k)
                || KEY_PM_AI_COLUMN_CONFIG_WORKBOOK.equals(k)
                || KEY_PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK.equals(k)
                || KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK.equals(k)
                || KEY_PM_AI_ACTUAL_DETAIL_WORKBOOK.equals(k);
    }

    /** {@link #KEY_PM_AI_PLAN_INPUT_PATH} (CSV / Parquet / Excel plan input). */
    public static boolean isPlanInputPathEnvKey(String key) {
        return key != null && KEY_PM_AI_PLAN_INPUT_PATH.equals(key.trim());
    }

    /** Master-data CSV / text paths ({@link #TABULAR_DATA_TABLE_PATH_KEYS}). */
    public static boolean isTabularDataTablePathEnvKey(String key) {
        return key != null && TABULAR_DATA_TABLE_PATH_KEYS.contains(key.trim());
    }

    /** Result-task column config CSV. */
    public static boolean isCsvFilePathEnvKey(String key) {
        return key != null && KEY_PM_AI_RESULT_TASK_COLUMN_CONFIG_CSV.equals(key.trim());
    }

    /**
     * {@code PM_AI_PYTHON} がディレクトリ（例: {@code pm-ai-data/runtime/python-embed}）のみを指しているとき、配下の
     * {@code python.exe} / {@code python3} / {@code python} に置き換える。{@link ProcessBuilder} は実行ファイルが必要で、
     * フォルダパスだと Windows でアクセス拒否（CreateProcess error=5）になる。
     *
     * @return 実行ファイルの絶対パス。フォルダだがインタプリタが無いときは空（呼び出し側で既定へフォールバック）。
     */
    public static String normalizePmAiPythonExecutable(String raw) {
        if (raw == null || raw.isBlank()) {
            return "";
        }
        String trimmed = raw.strip();
        Path p;
        try {
            p = Path.of(trimmed);
        } catch (InvalidPathException e) {
            return trimmed;
        }
        try {
            if (Files.isDirectory(p)) {
                for (String leaf : List.of("python.exe", "python3", "python")) {
                    Path cand = p.resolve(leaf);
                    if (Files.isRegularFile(cand)) {
                        return cand.toAbsolutePath().normalize().toString();
                    }
                }
                return "";
            }
        } catch (SecurityException e) {
            return trimmed;
        }
        return trimmed;
    }

    /**
     * ポータブル同梱の Python embed（{@code pm-ai-data/runtime/python-embed/python.exe}）を {@code start}
     * から親ディレクトリへ最大 {@value #BUNDLED_ANCHOR_WALK_MAX_PARENT_HOPS} 段まで辿って探す。
     *
     * <p>ショートカット起動などで {@code user.dir} がインストール根の直下でない場合でも検出できるようにする。
     *
     * @return 見つかったときは正規化済み絶対パス
     */
    public static Optional<Path> findPortablePythonEmbedExecutable(Path start) {
        if (start == null) {
            return Optional.empty();
        }
        Path cur = start.toAbsolutePath().normalize();
        for (int i = 0; i < BUNDLED_ANCHOR_WALK_MAX_PARENT_HOPS; i++) {
            Path exe =
                    cur.resolve("pm-ai-data")
                            .resolve("runtime")
                            .resolve("python-embed")
                            .resolve("python.exe");
            if (Files.isRegularFile(exe)) {
                return Optional.of(exe.toAbsolutePath().normalize());
            }
            Path parent = cur.getParent();
            if (parent == null || Objects.equals(parent, cur)) {
                break;
            }
            cur = parent;
        }
        return Optional.empty();
    }

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
                    return p;
                }
            }
            for (Path p : new Path[] {underRepo, underNested}) {
                if (Files.isDirectory(p)) {
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
        return Path.of(DEFAULT_PM_AI_TASK_INPUT_SOURCE_DIR);
    }

    /** Machining actual-detail export folder; optional {@code PM_AI_ACTUAL_DETAIL_SOURCE_DIR} in {@code ui}. */
    public static Path resolveActualDetailSourceDir(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String override = trim(u.get(KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR));
        if (!override.isEmpty()) {
            return Path.of(override).toAbsolutePath().normalize();
        }
        return Path.of(DEFAULT_PM_AI_ACTUAL_DETAIL_SOURCE_DIR);
    }

    /**
     * {@link #KEY_PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES} を解決する。不正な値は {@link #DEFAULT_PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES}
     * にフォールバック。0 以下は「上限なし」。
     */
    public static long resolveActualDetailRawMaxBytes(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String raw = trim(u.get(KEY_PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES));
        if (raw.isEmpty()) {
            return DEFAULT_PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES;
        }
        long parsed = parseEnvByteCountToLong(raw);
        if (parsed < 0) {
            return DEFAULT_PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES;
        }
        return parsed;
    }

    /**
     * 加工実績元ファイルが上限を超えるとき {@link IOException} を送出する。上限が 0 以下のときは何もしない。
     *
     * @param file 実ファイル（通常は {@link Files#isRegularFile(Path, java.nio.file.LinkOption...)}）
     */
    public static void ensureActualDetailRawFileWithinLimit(Path file, Map<String, String> ui)
            throws IOException {
        long max = resolveActualDetailRawMaxBytes(ui);
        if (max <= 0) {
            return;
        }
        if (file == null || !Files.isRegularFile(file)) {
            return;
        }
        long sz = Files.size(file);
        if (sz > max) {
            throw new IOException(
                    "加工実績の元データが大きすぎます（"
                            + sz
                            + " バイト）。上限 "
                            + max
                            + " バイト（環境変数 "
                            + KEY_PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES
                            + "）。値を引き上げるか、出力ファイルを分割してください。");
        }
    }

    /**
     * 環境変数のバイト数指定を解析する。{@code 20971520}、{@code 20M} / {@code 20MB}、{@code 8192K} 等。
     *
     * @return バイト数。0 は上限なし。「上限なし」は {@link #resolveActualDetailRawMaxBytes} がそのまま返す。
     *     負値は解析失敗。
     */
    static long parseEnvByteCountToLong(String raw) {
        if (raw == null) {
            return -1;
        }
        String s = raw.strip().replace("_", "").replace(" ", "");
        if (s.isEmpty()) {
            return -1;
        }
        String upper = s.toUpperCase(Locale.ROOT);
        long multiplier = 1;
        if (upper.endsWith("GB")) {
            multiplier = 1024L * 1024 * 1024;
            s = s.substring(0, s.length() - 2).strip();
        } else if (upper.endsWith("MB")) {
            multiplier = 1024L * 1024;
            s = s.substring(0, s.length() - 2).strip();
        } else if (upper.endsWith("KB")) {
            multiplier = 1024L;
            s = s.substring(0, s.length() - 2).strip();
        } else if (upper.endsWith("G")) {
            multiplier = 1024L * 1024 * 1024;
            s = s.substring(0, s.length() - 1).strip();
        } else if (upper.endsWith("M")) {
            multiplier = 1024L * 1024;
            s = s.substring(0, s.length() - 1).strip();
        } else if (upper.endsWith("K")) {
            multiplier = 1024L;
            s = s.substring(0, s.length() - 1).strip();
        }
        upper = s.toUpperCase(Locale.ROOT);
        if (upper.endsWith("B") && s.length() > 1) {
            char before = s.charAt(s.length() - 2);
            if (!Character.isDigit(before)) {
                s = s.substring(0, s.length() - 1).strip();
            }
        }
        try {
            long n = Long.parseLong(s);
            if (n == 0) {
                return 0;
            }
            return Math.multiplyExact(n, multiplier);
        } catch (NumberFormatException | ArithmeticException e) {
            return -1;
        }
    }

    /**
     * Directory for standalone result-dispatch xlsx; optional {@code PM_AI_RESULT_DISPATCH_TABLE_DIR} in
     * {@code ui}. Matches {@code planning_core.dispatch_workspace.resolve_result_dispatch_table_output_dir}:
     * optional override, then parent of {@link #KEY_PM_AI_PLAN_INPUT_PATH} when it is an existing Excel workbook,
     * else {@code resolveRepoRoot(ui)}/{@code code/output}.
     */
    public static Path resolveResultDispatchTableDir(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String override = trim(u.get(KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR));
        if (!override.isEmpty()) {
            return Path.of(override).toAbsolutePath().normalize();
        }
        String pip = trim(u.get(KEY_PM_AI_PLAN_INPUT_PATH));
        if (!pip.isEmpty()) {
            try {
                Path planInput = Path.of(pip);
                if (Files.isRegularFile(planInput)) {
                    String pl = pip.toLowerCase(Locale.ROOT);
                    if (pl.endsWith(".xlsx")
                            || pl.endsWith(".xlsm")
                            || pl.endsWith(".xltx")
                            || pl.endsWith(".xltm")) {
                        Path parent = planInput.toAbsolutePath().normalize().getParent();
                        if (parent != null) {
                            return parent;
                        }
                    }
                }
            } catch (Exception ignored) {
                // fall through to default (same drive / Unicode paths)
            }
        }
        return resolveRepoRoot(u).resolve("code").resolve("output").toAbsolutePath().normalize();
    }

    /** Basename of the JSON export for the result dispatch table (next to the standalone xlsx). */
    public static final String RESULT_DISPATCH_TABLE_JSON_BASENAME =
            "結果_配台表.json";

    /**
     * {@link #RESULT_DISPATCH_TABLE_JSON_BASENAME} under {@link #resolveResultDispatchTableDir(Map)} (override via
     * {@link #KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR}).
     */
    public static Path resolveResultDispatchTableJsonPath(Map<String, String> ui) {
        return resolveResultDispatchTableDir(ui != null ? ui : Map.of())
                .resolve(RESULT_DISPATCH_TABLE_JSON_BASENAME)
                .toAbsolutePath()
                .normalize();
    }

    /** Basename for the shaped Aladdin-plan cache JSON (colocated with the dispatch JSON). */
    public static final String SHAPED_ALADDIN_PLAN_JSON_BASENAME = "shaped_aladdin_plan.json";

    /** Basename for the shaped processing-actuals cache JSON (colocated with the dispatch JSON). */
    public static final String SHAPED_PROCESSING_ACTUALS_JSON_BASENAME =
            "shaped_processing_actuals.json";

    /** Path of {@link #SHAPED_ALADDIN_PLAN_JSON_BASENAME} next to the dispatch table JSON. */
    public static Path resolveShapedAladdinPlanJsonPath(Map<String, String> ui) {
        return resolveResultDispatchTableDir(ui != null ? ui : Map.of())
                .resolve(SHAPED_ALADDIN_PLAN_JSON_BASENAME)
                .toAbsolutePath()
                .normalize();
    }

    /** Path of {@link #SHAPED_PROCESSING_ACTUALS_JSON_BASENAME} next to the dispatch table JSON. */
    public static Path resolveShapedProcessingActualsJsonPath(Map<String, String> ui) {
        return resolveResultDispatchTableDir(ui != null ? ui : Map.of())
                .resolve(SHAPED_PROCESSING_ACTUALS_JSON_BASENAME)
                .toAbsolutePath()
                .normalize();
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
     * if set and a directory, else parent of the main-run macro-book path when provided, else parent of {@link
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

    /**
     * Stage1/2 の既定出力ディレクトリ（{@link #KEY_PM_AI_OUTPUT_DIR} または {@link #resolveRepoRoot(Map)} の直下
     * {@code output}）。
     */
    public static Path resolveDefaultOutputDir(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String override = trim(u.get(KEY_PM_AI_OUTPUT_DIR));
        if (!override.isEmpty()) {
            return Path.of(override).toAbsolutePath().normalize();
        }
        return resolveRepoRoot(u).resolve("output").toAbsolutePath().normalize();
    }

    /**
     * {@code code/} 配下のサマリ用マクロブック（{@code サマリ_AI配台.xlsm}）。実行・ログタブの「開く」から参照。
     */
    public static final String SUMMARY_AI_DISPATCH_XLSM =
            "サマリ_AI配台.xlsm";

    /**
     * リポジトリ {@code code/} 内の {@link #SUMMARY_AI_DISPATCH_XLSM} の絶対パス（{@link #resolveRepoRoot} と同一のルート解決）。
     * {@link #KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK} が非空のときはそのパス（絶対、または {@code code/} 基準の相対）を返す。
     */
    public static Path summaryAiDispatchXlsmPath(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String override = trim(u.get(KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK));
        if (!override.isEmpty()) {
            Path p = Path.of(override);
            if (!p.isAbsolute()) {
                p = resolveRepoRoot(u).resolve("code").resolve(override);
            }
            return p.toAbsolutePath().normalize();
        }
        return resolveRepoRoot(u)
                .resolve("code")
                .resolve(SUMMARY_AI_DISPATCH_XLSM)
                .toAbsolutePath()
                .normalize();
    }

    /** Filename for stage-1 shaped tasks ({@code planning_core.STAGE1_OUTPUT_FILENAME}). */
    public static final String STAGE1_PLAN_TASKS_FILENAME = "plan_input_tasks.xlsx";

    /** Sheet name in {@link #STAGE1_PLAN_TASKS_FILENAME} ({@code planning_core.run_stage1_extract} / {@code to_excel}). */
    public static final String STAGE1_PLAN_OUTPUT_SHEET = "タスク一覧";

    /**
     * Preview workbook written right after {@code load_tasks_df} ({@code planning_core.STAGE1_TASK_INPUT_PREVIEW_FILENAME}).
     */
    public static final String STAGE1_TASK_INPUT_PREVIEW_FILENAME = "stage1_task_input_table.xlsx";

    /** Sheet name inside {@link #STAGE1_TASK_INPUT_PREVIEW_FILENAME}. */
    public static final String STAGE1_TASK_INPUT_PREVIEW_SHEET = "タスク入力整形";

    /**
     * Written by {@code run_stage1_extract} beside {@code json_data_dir} ({@code planning_core} /
     * {@code STAGE1_EXCLUDE_RULES_JSON_FILENAME}).
     */
    public static final String STAGE1_EXCLUDE_RULES_JSON_FILENAME = "stage1_exclude_rules.json";

    /**
     * Path to the stage-1 exclude-rules sidecar JSON (same as Python {@code planning_core.bootstrap}
     * {@code json_data_dir}: {@code <effective cwd>/json/}, typically beside {@code output/} under {@code code/}).
     */
    public static Path stage1ExcludeRulesJsonPath(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path pyDir = resolvePythonScriptDir(u);
        Path codeDir = pyDir.getParent();
        Path underCodeJson =
                codeDir != null
                        ? codeDir.resolve("json").resolve(STAGE1_EXCLUDE_RULES_JSON_FILENAME)
                        : pyDir.resolve("json").resolve(STAGE1_EXCLUDE_RULES_JSON_FILENAME);
        return underCodeJson.toAbsolutePath().normalize();
    }

    /**
     * Legacy location used before aligning with Python {@code cwd/json}; checked if {@link #stage1ExcludeRulesJsonPath}
     * is missing.
     */
    public static Path stage1ExcludeRulesJsonPathLegacyUnderPython(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        return resolvePythonScriptDir(u)
                .resolve("json")
                .resolve(STAGE1_EXCLUDE_RULES_JSON_FILENAME)
                .toAbsolutePath()
                .normalize();
    }

    /**
     * Default for {@link #KEY_PM_AI_EXCLUDE_RULES_JSON}: {@code code/exclude_rules.json} when present, else
     * {@code code/json/stage1_exclude_rules.json} when present (repository typically ships the latter).
     */
    public static Optional<Path> resolveDefaultExcludeRulesJsonPath(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path primary = resolveRepoRoot(u).resolve("code").resolve("exclude_rules.json");
        if (Files.isRegularFile(primary)) {
            return Optional.of(primary.toAbsolutePath().normalize());
        }
        Path stage1 = stage1ExcludeRulesJsonPath(u);
        if (Files.isRegularFile(stage1)) {
            return Optional.of(stage1.toAbsolutePath().normalize());
        }
        return Optional.empty();
    }

    /**
     * Default path to stage-1 Excel output.
     *
     * <p>{@code planning_core.bootstrap} resolves {@code output_dir} from {@code PM_AI_OUTPUT_DIR} or
     * repository-root {@code output/} (see Python bootstrap). Legacy layouts under {@code code/output/} are still
     * detected when present.
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
        Path primary = resolveDefaultOutputDir(u).resolve(STAGE1_PLAN_TASKS_FILENAME);
        if (Files.isRegularFile(primary)) {
            return primary.toAbsolutePath().normalize();
        }
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
        return primary.toAbsolutePath().normalize();
    }

    /**
     * Directory where stage-2 writes {@code 計画*.xlsx} and {@code 人員*.xlsx}
     * (same folder as {@link #defaultStage1PlanTasksPath} — typically {@code .../code/output/}).
     */
    public static Path defaultPlanningOutputDir(Map<String, String> ui) {
        Path planTasks = defaultStage1PlanTasksPath(ui);
        Path parent = planTasks.getParent();
        if (parent != null) {
            return parent.toAbsolutePath().normalize();
        }
        return resolveDefaultOutputDir(ui != null ? ui : Map.of());
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
        Path primary = resolveDefaultOutputDir(u).resolve(STAGE1_TASK_INPUT_PREVIEW_FILENAME);
        if (Files.isRegularFile(primary)) {
            return primary.toAbsolutePath().normalize();
        }
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
        return primary.toAbsolutePath().normalize();
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

    /** {@link #resolveRepoRoot(Map)}/{@link #SPECIAL_RULES_SUMMARY_MD} */
    public static Path resolveSpecialRulesSummaryMd(Map<String, String> ui) {
        return resolveRepoRoot(ui).resolve(SPECIAL_RULES_SUMMARY_MD).toAbsolutePath().normalize();
    }

    /** {@link #resolveRepoRoot(Map)}/{@link #SPECIAL_RULES_ENUMERATED_MD} */
    public static Path resolveSpecialRulesEnumeratedMd(Map<String, String> ui) {
        return resolveRepoRoot(ui).resolve(SPECIAL_RULES_ENUMERATED_MD).toAbsolutePath().normalize();
    }

    /**
     * Discovers a macro {@code .xlsm} for auto-fill (JavaFX main-run tab field). Uses {@code PM_AI_WORKSPACE}
     * then {@link #resolveRepoRoot(Map)} scan. Not tied to an env-tab variable.
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
     * フォルダ系環境変数の値を、現在のリポジトリ根に対して補正できるときだけ置き換え文字列を返す。
     *
     * <ul>
     *   <li>{@link #KEY_PM_AI_REPO_ROOT}: 相対パスは {@link Path#toAbsolutePath()} で絶対化</li>
     *   <li>その他フォルダキー: リポジトリからの相対パスは {@link #resolveRepoRoot(Map)} に対して解決</li>
     *   <li>絶対パスが現在のリポジトリ配下なら正規化のみ</li>
     *   <li>別ルートにあった旧クローンの絶対パスは、パス内の {@link Path#getFileName() リポジトリ終端名}
     *       と一致する区切り以降を現在のリポジトリ根に再接続（サブパスのみ）</li>
     * </ul>
     *
     * <p>{@link #KEY_PM_AI_TASK_INPUT_SOURCE_DIR} / {@link #KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR} はネットワークソース正本のため常に空を返す。
     *
     * リポジトリ外を意図した相対パス（解決結果がリポジトリ根の外）は変更しない。
     */
    public static Optional<String> normalizeFolderEnvValue(Map<String, String> ui, String key, String rawValue) {
        String k = key != null ? key.trim() : "";
        if (!isFolderPathEnvKey(k)) {
            return Optional.empty();
        }
        if (KEY_PM_AI_TASK_INPUT_SOURCE_DIR.equals(k) || KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR.equals(k)) {
            return Optional.empty();
        }
        String v = trim(rawValue);
        if (v.isEmpty()) {
            return Optional.empty();
        }
        Path Rn = resolveRepoRoot(ui != null ? ui : Map.of()).toAbsolutePath().normalize();

        if (KEY_PM_AI_REPO_ROOT.equals(k)) {
            Path p = Path.of(v);
            Path out = p.isAbsolute() ? p.normalize() : p.toAbsolutePath().normalize();
            return pathsEqualString(v, out) ? Optional.empty() : Optional.of(out.toString());
        }

        Path p = Path.of(v);
        Path resolved;
        if (p.isAbsolute()) {
            Path pn = p.toAbsolutePath().normalize();
            if (isStrictlyUnderOrEqualRepo(pn, Rn)) {
                resolved = pn;
            } else {
                Path relocated = relocateUnderRepoByLeafName(pn, Rn);
                if (relocated != null && isStrictlyUnderOrEqualRepo(relocated, Rn)) {
                    resolved = relocated;
                } else {
                    return Optional.empty();
                }
            }
        } else {
            Path relResolved = Rn.resolve(p).normalize();
            if (!isStrictlyUnderOrEqualRepo(relResolved, Rn)) {
                return Optional.empty();
            }
            resolved = relResolved;
        }
        return pathsEqualString(v, resolved) ? Optional.empty() : Optional.of(resolved.toString());
    }

    /**
     * {@code ui} のフォルダ系キーを {@link #FOLDER_PATH_NORMALIZE_ORDER} の順で更新した差分（キー→新値）。
     * 途中で {@link #KEY_PM_AI_REPO_ROOT} が変わると後続キーの解決に反映される。
     */
    public static Map<String, String> normalizedFolderEnvOverrides(Map<String, String> ui) {
        Map<String, String> work = new HashMap<>(ui != null ? ui : Map.of());
        Map<String, String> overrides = new LinkedHashMap<>();
        for (String fk : FOLDER_PATH_NORMALIZE_ORDER) {
            String raw = trim(work.get(fk));
            Optional<String> n = normalizeFolderEnvValue(work, fk, raw);
            if (n.isPresent()) {
                String nv = n.get();
                overrides.put(fk, nv);
                work.put(fk, nv);
            }
        }
        return overrides;
    }

    private static boolean isStrictlyUnderOrEqualRepo(Path path, Path repoNorm) {
        Path pn = path.toAbsolutePath().normalize();
        Path rn = repoNorm.toAbsolutePath().normalize();
        return pn.startsWith(rn);
    }

    /**
     * {@code absoluteForeign} の祖先に {@code repoNorm.getFileName()} と同名の区切りがあれば、その直下を {@code repoNorm}
     * に付け替えたパスを返す。
     */
    static Path relocateUnderRepoByLeafName(Path absoluteForeign, Path repoNorm) {
        Path rn = repoNorm.toAbsolutePath().normalize();
        Path leaf = rn.getFileName();
        if (leaf == null) {
            return null;
        }
        String marker = leaf.toString();
        Path pn = absoluteForeign.toAbsolutePath().normalize();
        int n = pn.getNameCount();
        for (int i = 0; i < n; i++) {
            if (marker.equals(pn.getName(i).toString())) {
                if (i + 1 >= n) {
                    return rn;
                }
                Path tail = pn.subpath(i + 1, n);
                return rn.resolve(tail).normalize();
            }
        }
        return null;
    }

    private static boolean pathsEqualString(String rawTrimmed, Path resolved) {
        Path before = Path.of(rawTrimmed);
        Path bNorm = before.isAbsolute() ? before.normalize() : before.toAbsolutePath().normalize();
        return bNorm.equals(resolved.toAbsolutePath().normalize());
    }

    /**
     * Lists {@code .xlsm} in a directory; if one file, returns it; if several, prefers a name
     * containing {@code 配台}, else lexicographically first.
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
        String marker = "配台";
        Optional<Path> preferred = xlsms.stream()
                .filter(p -> p.getFileName().toString().contains(marker))
                .min(Comparator.comparing(p -> p.getFileName().toString()));
        return preferred.or(() -> xlsms.stream()
                .min(Comparator.comparing(p -> p.getFileName().toString())));
    }

    private static Optional<Path> findCodePythonFrom(Path start) {
        Path cur = start;
        for (int i = 0; i < BUNDLED_ANCHOR_WALK_MAX_PARENT_HOPS; i++) {
            Path candidate = cur.resolve("code").resolve("python");
            if (Files.isDirectory(candidate) && Files.isRegularFile(candidate.resolve("task_extract_stage1.py"))) {
                return Optional.of(candidate.toAbsolutePath().normalize());
            }
            Path bundled =
                    cur.resolve("pm-ai-data").resolve("code").resolve("python");
            if (Files.isDirectory(bundled) && Files.isRegularFile(bundled.resolve("task_extract_stage1.py"))) {
                return Optional.of(bundled.toAbsolutePath().normalize());
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
