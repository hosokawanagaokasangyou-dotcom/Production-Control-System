package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.InvalidPathException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
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

    /**
     * 材料テーブル等 {@code code/} 配下 CSV の正本ディレクトリ。段階1子プロセスへ明示渡しし、Python と Java UI の読み書き先を揃える。
     */
    public static final String KEY_PM_AI_CODE_DIR = "PM_AI_CODE_DIR";

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
     * 配台手動修正: 段階3試行後の日付セルで {@code （段階3前）（段階3後）} を改行せず1行（スペース区切り）表示する。
     * {@code 1}/{@code true}/{@code yes}/{@code on} で有効。未設定または {@code 0}/{@code false} 等で従来の2行表示。
     */
    public static final String KEY_PM_AI_DEBUG_STAGE3_PLAN_ACTUAL_SINGLE_LINE =
            "PM_AI_DEBUG_STAGE3_PLAN_ACTUAL_SINGLE_LINE";

    /**
     * Stage1/2 成果物フォルダ（従来の {@code code/output} に相当）。未設定時は {@link #resolveRepoRoot(Map)} の直下の
     * {@code output/}。Python {@code planning_core.bootstrap} の {@code output_dir} と揃える。
     */
    public static final String KEY_PM_AI_OUTPUT_DIR = "PM_AI_OUTPUT_DIR";

    /**
     * 配台システム起動時に選択した操作者名（{@link FactoryOperatorUserStore}）。子プロセス env に載せる。
     */
    public static final String KEY_PM_AI_OPERATOR_USER = "PM_AI_OPERATOR_USER";

    public static final String KEY_PM_AI_TASK_INPUT_SOURCE_DIR = "PM_AI_TASK_INPUT_SOURCE_DIR";

    /** Folder for machining actual-detail Excel exports (PQ plan/02 {@code Folder.Files}). */
    public static final String KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR = "PM_AI_ACTUAL_DETAIL_SOURCE_DIR";

    /**
     * 依頼書入力（アラジンマスタ）フォルダ。後加工商品／加工内容／工程マスタおよび
     * {@code マスタリレーション統合結果.xlsx} を置くディレクトリ（フルパス）。
     */
    public static final String KEY_PM_AI_ALADDIN_MASTER_DIR = "PM_AI_ALADDIN_MASTER_DIR";

    /**
     * 後加工商品マスタのアップロード用 xlsx（フルパス）。空のときは
     * {@link #resolveAladdinMasterDir(Map)}/{@link jp.co.pm.ai.desktop.io.PostProcessingProductMasterIo#DEFAULT_UPLOAD_FILE_NAME}。
     */
    public static final String KEY_PM_AI_POSTPROC_PRODUCT_MASTER_UPLOAD =
            "PM_AI_POSTPROC_PRODUCT_MASTER_UPLOAD";

    /** {@link #KEY_PM_AI_ALADDIN_MASTER_DIR} 未設定時、{@link #resolveRepoRoot(Map)} 直下のサブフォルダ名。 */
    public static final String ALADDIN_MASTER_DIR_LEAF_NAME = "アラジンマスタ";

    /**
     * 依頼書入力の受注データベース Excel（受注ﾌｧｲﾙ シートを含むブック）のフルパス。
     */
    public static final String KEY_PM_AI_REQUEST_FORM_JUCHU_FILE = "PM_AI_REQUEST_FORM_JUCHU_FILE";

    /**
     * 依頼書入力がスキャンする依頼書原本フォルダ（{@code *加工依頼書*.xlsm} 等を含むディレクトリ）。
     */
    public static final String KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR = "PM_AI_REQUEST_FORM_ORIGINAL_DIR";

    /**
     * 依頼書入力タブ「リモートデスクトップ」で起動する RDP プロファイル（{@code *.rdp}）のフルパス。
     */
    public static final String KEY_PM_AI_REQUEST_FORM_RDP_PROFILE = "PM_AI_REQUEST_FORM_RDP_PROFILE";

    /**
     * 接続先サーバー上で RDP 接続時に起動するプログラム（.rdp の alternate shell（RemoteApp ではない））。空なら無効。
     */
    public static final String KEY_PM_AI_RDP_COMPANION_PROGRAM = "PM_AI_RDP_COMPANION_PROGRAM";

    /** {@link #KEY_PM_AI_RDP_COMPANION_PROGRAM} の引数（alternate shell へ付与する引数）。 */
    public static final String KEY_PM_AI_RDP_COMPANION_PROGRAM_ARGS = "PM_AI_RDP_COMPANION_PROGRAM_ARGS";

    /** 接続先 RDP ランチャー exe（{@link #RDP_LAUNCHER_EXE_BASENAME}）のフルパス上書き。 */
    public static final String KEY_PM_AI_RDP_LAUNCHER_EXE = "PM_AI_RDP_LAUNCHER_EXE";

    /** 接続先 RDP ランチャー設定 ini（{@link #RDP_LAUNCHER_INI_BASENAME}）のフルパス上書き。 */
    public static final String KEY_PM_AI_RDP_LAUNCHER_INI = "PM_AI_RDP_LAUNCHER_INI";

    /** {@code 1} のときのみ .rdp へ alternate shell（リモート起動プログラム）を書込。 */
    public static final String KEY_PM_AI_RDP_EMBED_STARTUP_IN_PROFILE = "PM_AI_RDP_EMBED_STARTUP_IN_PROFILE";

    /** 接続先ランチャー exe の自動再配備（{@code 0/false/off} で無効）。 */
    public static final String KEY_PM_AI_RDP_LAUNCHER_AUTO_DEPLOY = "PM_AI_RDP_LAUNCHER_AUTO_DEPLOY";

    public static final String RDP_LAUNCHER_EXE_BASENAME = "PmAiRdpRemoteLauncher.exe";
    public static final String RDP_LAUNCHER_VERSION_BASENAME = "PmAiRdpRemoteLauncher.version.txt";
    public static final String RDP_LAUNCHER_INI_BASENAME = "RAP設定.ini";

    /**
     * 依頼書 PDF プレビュー: Type0 日本語フォントのサイズ補正係数（Excel pt に乗算）。{@code 0.50}～{@code 1.00}。
     */
    public static final String KEY_PM_AI_REQUEST_FORM_PREVIEW_PDF_CJK_SCALE =
            "PM_AI_REQUEST_FORM_PREVIEW_PDF_CJK_SCALE";

    /** {@link #KEY_PM_AI_REQUEST_FORM_PREVIEW_PDF_CJK_SCALE} の既定（はみ出し抑制）。 */
    public static final float DEFAULT_PM_AI_REQUEST_FORM_PREVIEW_PDF_CJK_SCALE = 0.72f;

    /** {@link #KEY_PM_AI_REQUEST_FORM_JUCHU_FILE} 未設定時のファイル名（作業フォルダ直下）。 */
    public static final String DEFAULT_REQUEST_FORM_JUCHU_FILE_NAME = "加工依頼書入力.xlsm";

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

    /** §B 特別ルール DSL JSON（{@link jp.co.pm.ai.desktop.dispatch.rules.paths.DispatchRulePaths}）。 */
    public static final String KEY_PM_AI_DISPATCH_SPECIAL_RULES_JSON =
            jp.co.pm.ai.desktop.dispatch.rules.paths.DispatchRulePaths
                    .KEY_PM_AI_DISPATCH_SPECIAL_RULES_JSON;

    public static final String KEY_PM_AI_DISPATCH_RULE_ENGINE =
            jp.co.pm.ai.desktop.dispatch.rules.paths.DispatchRulePaths.KEY_PM_AI_DISPATCH_RULE_ENGINE;

    public static final String KEY_PM_AI_DISPATCH_RULE_LEGACY_FALLBACK =
            jp.co.pm.ai.desktop.dispatch.rules.paths.DispatchRulePaths
                    .KEY_PM_AI_DISPATCH_RULE_LEGACY_FALLBACK;

    /** リポジトリ直下の人間向け要約（デスクトップ「特別ルール」タブと運用で同期）。 */
    public static final String SPECIAL_RULES_SUMMARY_MD = "特別ルール.md";

    /** リポジトリ直下の L 番号列挙（{@code planning_core/_core.py} のコメントと対応）。 */
    public static final String SPECIAL_RULES_ENUMERATED_MD = "特別ルール列挙.md";

    /** 取扱説明書 HTML の相対パス（{@link #resolveManualIndexHtml(Map)}）。 */
    public static final String MANUAL_INDEX_HTML_REL = "manual/html/index.html";

    /** 配台業務ルール HTML（{@link #resolveDispatchRulesHtml(Map)}）。 */
    public static final String DISPATCH_RULES_HTML_REL = "code/要件定義/配台ルール.html";

    /** 現場向け Word 手順書（リポジトリ直下・{@link #resolveDispatchUsageGuideDocx(Map)}）。 */
    public static final String DISPATCH_USAGE_GUIDE_DOCX = "配台システム使い方（整理版）.docx";

    /** マスタブック（{@code master.xlsm} 等）の絶対パスまたは {@code code/} 相対。planning_core 子プロセスの必須 env。 */
    public static final String KEY_PM_AI_MASTER_WORKBOOK = "PM_AI_MASTER_WORKBOOK";

    /**
     * @deprecated {@link #KEY_PM_AI_MASTER_WORKBOOK} に一本化。セッション移行・除去用のキー名のみ。
     */
    @Deprecated
    public static final String KEY_MASTER_WORKBOOK_FILE = "MASTER_WORKBOOK_FILE";

    /**
     * サマリ AI 配台 Excel（{@link #SUMMARY_AI_DISPATCH_XLSX} 等）の絶対パス、または {@code code/} 相対。
     * 空で {@code code/} 既定名。親フォルダは工場共有 DATA の根（操作者 bin、配台除外 JSON、ルックアップ CSV、
     * 実行時間履歴、設備ガント PDF、サマリ世代退避等の sibling 解決基準）。{@link #summaryAiDispatchXlsxPathForFactory}
     * で利用工場と不一致のとき工場既定 UNC へ切替。
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

    /** 1 のときデータ抽出日（当日）の配台を行わず、翌暦日以降を計画開始日とする（段階2）。UI は配台計画_タスク入力タブ。 */
    public static final String KEY_PM_AI_STAGE2_SKIP_TODAY_DISPATCH = "PM_AI_STAGE2_SKIP_TODAY_DISPATCH";

    /**
     * 1 のとき実加工数が正の行（加工途中相当）を配台キューに載せない（当日完了と想定、段階2）。JavaFX 段階2 では常に 0（翌日配台量はタスク入力タブ）。
     */
    public static final String KEY_PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH =
            "PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH";

    /**
     * 段階2: 加工途中行の翌日配台量 (m) を載せた UTF-8 JSON の絶対パス。JavaFX が段階2直前に書き、{@code build_task_queue_from_planning_df}
     * が実加工数&gt;0 の行の配台量を上書きする。無効・未設定時はシートの配台使用残数量を用いる。
     */
    public static final String KEY_PM_AI_STAGE2_IN_PROGRESS_NEXT_DAY_DISPATCH_JSON =
            "PM_AI_STAGE2_IN_PROGRESS_NEXT_DAY_DISPATCH_JSON";

    /**
     * 段階2.1 残業シミュレーション: JavaFX が書く UTF-8 JSON（working_overrides / overtime_minutes）。
     * 段階2.1 子プロセス起動時のみ載せ、master.xlsm は変更しない。
     */
    public static final String KEY_PM_AI_OVERTIME_SIMULATION_JSON =
            "PM_AI_OVERTIME_SIMULATION_JSON";

    /** 段階2.1: 残業/休出シミュのフル再配台（段階2 成果物は上書きしない）。 */
    public static final String KEY_PM_AI_STAGE2_1_OVERTIME = "PM_AI_STAGE2_1_OVERTIME";

    /** {@code output/stage21/} 配下の段階2.1 成果物ディレクトリ名。 */
    public static final String STAGE21_OUTPUT_SUBDIR = "stage21";

    /**
     * 段階2の実行エンジン（互換用キー）。JavaFX 実行タブから段階2を起動するときは常に Python 子プロセス（{@code
     * plan_simulation_stage2.py}）のみ。未設定・空・{@code python}（大小無視）で従来どおり。{@code java} が指定されていても無視され Python
     * が起動する（旧 JVM 段階2は撤去済み）。
     */
    public static final String KEY_PM_AI_STAGE2_ENGINE = "PM_AI_STAGE2_ENGINE";

    /**
     * 段階2の Excel 成果物（結果ブック）のフォントファミリ。空のときは planning_core の {@code RESULT_BOOK_FONT_NAME}（BIZ
     * UDゴシック）相当。JavaFX 実行タブのコンボで上書き可。
     */
    public static final String KEY_PM_AI_RESULT_BOOK_FONT = "PM_AI_RESULT_BOOK_FONT";

    /**
     * 1/true/yes/on のとき planning_core / JavaFX から Gemini {@code generateContent} を呼ばない（開発用）。
     * JavaFX 実行・ログタブのチェックが子プロセス起動時に上書きする。
     */
    public static final String KEY_PM_AI_SKIP_GEMINI_API = "PM_AI_SKIP_GEMINI_API";

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
     * 湖南工場共有の {@code pm-ai-package-release} フォルダ（UNC）。{@link #KEY_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR} の湖南既定。
     * 直下に外付け {@code version.txt} と {@code PMD_version_upgrade.zip} を置く。
     */
    public static final String DEFAULT_PM_AI_PORTABLE_BUNDLE_RELEASE_DIR_KONAN =
            "\\\\192.168.0.101\\共有フォルダ\\湖南工場\\湖南共有\\002  加工G\\●配台AIシステム\\pm-ai-package-release";

    /**
     * 国分工場共有の {@code pm-ai-package-release} フォルダ（UNC）。{@link FactorySite#KOKUBU} のバージョンアップ正本。
     */
    public static final String DEFAULT_PM_AI_PORTABLE_BUNDLE_RELEASE_DIR_KOKUBU =
            "\\\\192.168.0.101\\共有フォルダ\\国分工場\\国分共有\\●配台AIシステム\\pm-ai-package-release";

    /**
     * 環境変数タブで値が空のときの正本（UNC）。{@link #DEFAULT_PM_AI_PORTABLE_BUNDLE_RELEASE_DIR_KONAN} と同じ（湖南）。
     * ZIP フルパス指定も可だが、運用の正は release フォルダパス。
     */
    public static final String DEFAULT_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR =
            DEFAULT_PM_AI_PORTABLE_BUNDLE_RELEASE_DIR_KONAN;

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
            KEY_PM_AI_ALADDIN_MASTER_DIR,
            KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR,
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
                    KEY_PM_AI_ALADDIN_MASTER_DIR,
                    KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR,
                    KEY_PM_AI_OUTPUT_DIR,
                    KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR,
                    KEY_COMPARE_GANTT_SNAPSHOT_DIR);

    /** Env keys whose value is a single file path (file chooser in the UI). */
    private static final Set<String> FILE_PATH_ENV_KEYS = createFilePathEnvKeys();

    private static Set<String> createFilePathEnvKeys() {
        HashSet<String> s = new HashSet<>();
        s.add(KEY_GEMINI_CREDENTIALS_JSON);
        s.add(KEY_PM_AI_EXCLUDE_RULES_JSON);
        s.add(KEY_PM_AI_DISPATCH_SPECIAL_RULES_JSON);
        s.add(KEY_PM_AI_MASTER_WORKBOOK);
        s.add(KEY_PM_AI_COLUMN_CONFIG_WORKBOOK);
        s.add(KEY_PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK);
        s.add(KEY_PM_AI_RESULT_TASK_COLUMN_CONFIG_CSV);
        s.add(KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK);
        s.add(KEY_PM_AI_PLAN_INPUT_PATH);
        s.add(KEY_PM_AI_PROCESSING_PLAN_PATH);
        s.add(KEY_PM_AI_ACTUAL_DETAIL_WORKBOOK);
        s.add(KEY_PM_AI_REQUEST_FORM_JUCHU_FILE);
        s.add(KEY_PM_AI_REQUEST_FORM_RDP_PROFILE);
        s.add(KEY_PM_AI_PLAN_RESULT_TASK_JSON_PATH);
        s.add(KEY_PM_AI_CURSOR_DEBUG_LOG);
        s.add(KEY_PM_AI_DEBUG_LOG_MIRROR);
        s.add(KEY_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR);
        s.addAll(TABULAR_DATA_TABLE_PATH_KEYS);
        return Set.copyOf(s);
    }

    private AppPaths() {}

    /**
     * 環境変数タブの値が truthy か（{@code 0}/{@code false}/{@code no}/{@code off}/{@code none} は false）。
     *
     * @param defaultWhenBlank キー未設定・空文字のときの戻り値
     */
    public static boolean isTruthyUiEnv(
            Map<String, String> ui, String key, boolean defaultWhenBlank) {
        if (key == null || key.isBlank()) {
            return defaultWhenBlank;
        }
        if (ui == null || ui.isEmpty()) {
            return defaultWhenBlank;
        }
        String v = ui.get(key);
        if (v == null || v.isBlank()) {
            return defaultWhenBlank;
        }
        String t = v.trim().toLowerCase(java.util.Locale.ROOT);
        return !java.util.List.of("0", "false", "no", "off", "none").contains(t);
    }

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
                || KEY_PM_AI_ACTUAL_DETAIL_WORKBOOK.equals(k)
                || KEY_PM_AI_REQUEST_FORM_JUCHU_FILE.equals(k);
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
     * {@code PM_AI_CODE_PYTHON_DIR} が pm-ai-data 配下を指していても、{@code PM_AI_REPO_ROOT/code/python}
     * が有効なら開発ツリー側を優先する（同梱 pm-ai-data が古いときの取りこぼし防止）。
     */
    public static Path resolvePythonScriptDir(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path repoPython = resolveRepoCodePythonDir(u);
        String override = trim(u.get(KEY_PM_AI_CODE_PYTHON_DIR));
        if (!override.isEmpty()) {
            Path p = Path.of(override).toAbsolutePath().normalize();
            if (Files.isDirectory(p)) {
                if (repoPython != null && isUnderPmAiDataCodePython(p)) {
                    try {
                        if (!Files.isSameFile(p, repoPython)) {
                            return repoPython;
                        }
                    } catch (IOException ignored) {
                        return repoPython;
                    }
                }
                return p;
            }
        }
        if (repoPython != null) {
            return repoPython;
        }
        String repo = trim(u.get(KEY_PM_AI_REPO_ROOT));
        if (!repo.isEmpty()) {
            Path base = Path.of(repo).toAbsolutePath().normalize();
            Path underNested =
                    base.resolve("Production-Control-System").resolve("code").resolve("python");
            if (Files.isDirectory(underNested)) {
                return underNested;
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

    /**
     * 材料テーブル CSV の正本 {@code code/} ディレクトリ。{@link #KEY_PM_AI_CODE_DIR} → {@link #resolvePythonScriptDir} の親
     * （{@code code}）→ {@link #resolveRepoRoot} の {@code code/} の順。
     */
    public static Path resolveCodeDir(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String override = trim(u.get(KEY_PM_AI_CODE_DIR));
        if (!override.isEmpty()) {
            Path p = Path.of(override).toAbsolutePath().normalize();
            if (Files.isDirectory(p)) {
                return p;
            }
        }
        Path py = resolvePythonScriptDir(u);
        Path parent = py.getParent();
        if (parent != null
                && "code".equals(parent.getFileName() != null ? parent.getFileName().toString() : "")
                && Files.isDirectory(parent)) {
            return parent.toAbsolutePath().normalize();
        }
        return resolveRepoRoot(u).resolve("code").toAbsolutePath().normalize();
    }

    /** {@code PM_AI_REPO_ROOT/code/python} が段階1スクリプトを含むときそのパス、無ければ {@code null}。 */
    static Path resolveRepoCodePythonDir(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String repo = trim(u.get(KEY_PM_AI_REPO_ROOT));
        if (repo.isEmpty()) {
            return null;
        }
        Path base = Path.of(repo).toAbsolutePath().normalize();
        Path underRepo = base.resolve("code").resolve("python");
        if (Files.isDirectory(underRepo)
                && Files.isRegularFile(underRepo.resolve("task_extract_stage1.py"))) {
            return underRepo;
        }
        Path underNested = base.resolve("Production-Control-System").resolve("code").resolve("python");
        if (Files.isDirectory(underNested)
                && Files.isRegularFile(underNested.resolve("task_extract_stage1.py"))) {
            return underNested;
        }
        return null;
    }

    static boolean isUnderPmAiDataCodePython(Path p) {
        if (p == null) {
            return false;
        }
        Path norm = p.toAbsolutePath().normalize();
        if (!"python".equals(norm.getFileName() != null ? norm.getFileName().toString() : "")) {
            return false;
        }
        Path code = norm.getParent();
        if (code == null || !"code".equals(code.getFileName() != null ? code.getFileName().toString() : "")) {
            return false;
        }
        Path pmAiData = code.getParent();
        return pmAiData != null
                && "pm-ai-data".equalsIgnoreCase(
                        pmAiData.getFileName() != null ? pmAiData.getFileName().toString() : "");
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
     * 依頼書入力向けアラジンマスタフォルダ。{@link #KEY_PM_AI_ALADDIN_MASTER_DIR} が空のときは
     * {@link GlobalInitSettingTarget#load()} の工場に応じた UNC（湖南／国分）。
     */
    public static Path resolveAladdinMasterDir(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String override = trim(u.get(KEY_PM_AI_ALADDIN_MASTER_DIR));
        if (!override.isEmpty()) {
            return Path.of(override).toAbsolutePath().normalize();
        }
        return Path.of(defaultAladdinMasterDirForFactory(GlobalInitSettingTarget.load()));
    }

    /** {@link FactorySite} 別の {@link #KEY_PM_AI_ALADDIN_MASTER_DIR} 既定 UNC。 */
    public static String defaultAladdinMasterDirForFactory(FactorySite site) {
        if (site == FactorySite.KOKUBU) {
            return DEFAULT_PM_AI_ALADDIN_MASTER_DIR_KOKUBU;
        }
        return DEFAULT_PM_AI_ALADDIN_MASTER_DIR_KONAN;
    }

    /**
     * 依頼書入力の受注ファイル Excel。{@link #KEY_PM_AI_REQUEST_FORM_JUCHU_FILE} が空のときは
     * {@link GlobalInitSettingTarget#load()} の工場に応じた UNC 既定。
     */
    public static Optional<Path> resolveRequestFormJuchuFile(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String override = trim(u.get(KEY_PM_AI_REQUEST_FORM_JUCHU_FILE));
        if (!override.isEmpty()) {
            return Optional.of(Path.of(override).toAbsolutePath().normalize());
        }
        return Optional.of(Path.of(defaultRequestFormJuchuFileForFactory(GlobalInitSettingTarget.load())));
    }

    /** {@link FactorySite} 別の {@link #KEY_PM_AI_REQUEST_FORM_JUCHU_FILE} 既定 UNC。 */
    public static String defaultRequestFormJuchuFileForFactory(FactorySite site) {
        if (site == FactorySite.KOKUBU) {
            return DEFAULT_PM_AI_REQUEST_FORM_JUCHU_FILE_KOKUBU;
        }
        return DEFAULT_PM_AI_REQUEST_FORM_JUCHU_FILE_KONAN;
    }

    /**
     * 依頼書原本フォルダ。{@link #KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR} が空のときは
     * {@link GlobalInitSettingTarget#load()} の工場に応じ、受注ファイル既定の親ディレクトリ。
     */
    public static Path resolveRequestFormOriginalDir(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String override = trim(u.get(KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR));
        if (!override.isEmpty()) {
            return Path.of(override).toAbsolutePath().normalize();
        }
        return Path.of(defaultRequestFormOriginalDirForFactory(GlobalInitSettingTarget.load()))
                .toAbsolutePath()
                .normalize();
    }

    /**
     * 依頼書入力タブの RDP プロファイル（{@code *.rdp}）。未設定・空は empty。
     */
    public static Optional<Path> resolveRequestFormRdpProfile(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String override = trim(u.get(KEY_PM_AI_REQUEST_FORM_RDP_PROFILE));
        if (override.isEmpty()) {
            return Optional.empty();
        }
        return Optional.of(Path.of(override).toAbsolutePath().normalize());
    }

    /** {@link FactorySite} 別の {@link #KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR} 既定（受注ファイルの親フォルダ）。 */
    public static String defaultRequestFormOriginalDirForFactory(FactorySite site) {
        Path parent = Path.of(defaultRequestFormJuchuFileForFactory(site)).getParent();
        return parent != null ? parent.toAbsolutePath().normalize().toString() : "";
    }

    /**
     * {@link #KEY_PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES} を解決する。不正な値は {@link #DEFAULT_PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES}
     * にフォールバック。0 以下は「上限なし」。
     */
    /**
     * {@link #KEY_PM_AI_REQUEST_FORM_PREVIEW_PDF_CJK_SCALE} を解決する。空・不正は {@link
     * #DEFAULT_PM_AI_REQUEST_FORM_PREVIEW_PDF_CJK_SCALE}。範囲外は {@code 0.50}～{@code 1.00} にクランプ。
     */
    public static float resolveRequestFormPreviewPdfCjkScale(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String raw = trim(u.get(KEY_PM_AI_REQUEST_FORM_PREVIEW_PDF_CJK_SCALE));
        if (raw.isEmpty()) {
            return DEFAULT_PM_AI_REQUEST_FORM_PREVIEW_PDF_CJK_SCALE;
        }
        try {
            float v = Float.parseFloat(raw.replace(',', '.'));
            if (Float.isNaN(v) || Float.isInfinite(v)) {
                return DEFAULT_PM_AI_REQUEST_FORM_PREVIEW_PDF_CJK_SCALE;
            }
            return Math.max(0.50f, Math.min(1.00f, v));
        } catch (NumberFormatException ex) {
            return DEFAULT_PM_AI_REQUEST_FORM_PREVIEW_PDF_CJK_SCALE;
        }
    }

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

    /** 段階2 出力の {@link #RESULT_DISPATCH_TABLE_JSON_BASENAME}。 */
    public static Path resolveResultDispatchTableStage2JsonPath(Map<String, String> ui) {
        return resolveResultDispatchTableDir(ui != null ? ui : Map.of())
                .resolve(RESULT_DISPATCH_TABLE_JSON_BASENAME)
                .toAbsolutePath()
                .normalize();
    }

    /** 段階2.1 残業シミュの出力ディレクトリ（段階2 の {@link #resolveResultDispatchTableDir} 配下）。 */
    public static Path resolveStage21OutputDir(Map<String, String> ui) {
        return resolveResultDispatchTableDir(ui != null ? ui : Map.of())
                .resolve(STAGE21_OUTPUT_SUBDIR)
                .toAbsolutePath()
                .normalize();
    }

    /** 段階2.1 の {@link #RESULT_DISPATCH_TABLE_JSON_BASENAME}。 */
    public static Path resolveStage21ResultDispatchJsonPath(Map<String, String> ui) {
        return resolveStage21OutputDir(ui != null ? ui : Map.of())
                .resolve(RESULT_DISPATCH_TABLE_JSON_BASENAME)
                .toAbsolutePath()
                .normalize();
    }

    /**
     * 手動修正・納期ビュー等が参照する配台表 JSON（段階2 出力の {@link #RESULT_DISPATCH_TABLE_JSON_BASENAME}）。
     *
     * <p>段階2 の Python 出力先も {@link #resolveResultDispatchTableStage2JsonPath(Map)} と同一。
     */
    public static Path resolveResultDispatchTableJsonPath(Map<String, String> ui) {
        return resolveResultDispatchTableStage2JsonPath(ui);
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
     * {@link #KEY_PM_AI_MASTER_WORKBOOK} を基準にマスタブックパスを解決する（planning_core の cwd 基準 relative と整合）。
     *
     * <p>PM_AI 未設定時は {@code master.xlsm} を {@link #resolveEffectivePlanningCwd} 基準の relative として解決する。
     */
    public static Path resolveMasterWorkbookPathResolved(Map<String, String> ui, String taskInputWorkbookPath) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String alt = trim(u.get(KEY_PM_AI_MASTER_WORKBOOK));
        if (!alt.isEmpty()) {
            if (alt.startsWith("\\\\")) {
                return Path.of(alt);
            }
            Path p = Path.of(alt);
            if (p.isAbsolute()) {
                return p.normalize();
            }
            Path cwd = resolveEffectivePlanningCwd(u, taskInputWorkbookPath);
            return cwd.resolve(alt).normalize().toAbsolutePath();
        }
        Path cwd = resolveEffectivePlanningCwd(u, taskInputWorkbookPath);
        return cwd.resolve("master.xlsm").normalize().toAbsolutePath();
    }

    /**
     * 廃止した {@link #KEY_MASTER_WORKBOOK_FILE} を {@link #KEY_PM_AI_MASTER_WORKBOOK} へ移行する値を返す。
     *
     * <p>PM_AI が既に非空のときは empty（上書きしない）。
     */
    public static Optional<String> migrateLegacyMasterWorkbookFileToPmAi(
            Map<String, String> ui, String legacyMasterWorkbookFile) {
        if (legacyMasterWorkbookFile == null || legacyMasterWorkbookFile.isBlank()) {
            return Optional.empty();
        }
        Map<String, String> u = ui != null ? new java.util.HashMap<>(ui) : new java.util.HashMap<>();
        if (!trim(u.get(KEY_PM_AI_MASTER_WORKBOOK)).isEmpty()) {
            return Optional.empty();
        }
        String legacy = legacyMasterWorkbookFile.trim();
        u.put(KEY_PM_AI_MASTER_WORKBOOK, legacy);
        Path resolved = resolveMasterWorkbookPathForDesktopOpen(u, "");
        if (Files.isRegularFile(resolved)) {
            return Optional.of(resolved.toAbsolutePath().normalize().toString());
        }
        if (legacy.startsWith("\\\\")) {
            return Optional.of(legacy);
        }
        Path leg = Path.of(legacy);
        if (leg.isAbsolute()) {
            return Optional.of(leg.normalize().toString());
        }
        return Optional.of(legacy);
    }

    /**
     * 実行・ログタブやマスタ読込サマリの「Excel を開く」用のマスタ解決。
     *
     * <p>{@link #resolveMasterWorkbookPathResolved} は planning_core と揃え {@link #resolveEffectivePlanningCwd} を
     * 基準にするため、段階2の production_plan が {@code output/} 配下にあると basename のマスタが
     * {@code output/国分master.xlsm} のように誤解決し得る。本メソッドはまず同一解決を試し、ファイルが無ければ
     * {@code code/python} の親（{@code code/}）、リポジトリ {@code code/}・{@code plan/}・ルートを順に探す。
     */
    public static Path resolveMasterWorkbookPathForDesktopOpen(Map<String, String> ui, String taskInputWorkbookPath) {
        Path primary = resolveMasterWorkbookPathResolved(ui, taskInputWorkbookPath);
        if (Files.isRegularFile(primary)) {
            return primary;
        }
        Map<String, String> u = ui != null ? ui : Map.of();
        String alt = trim(u.get(KEY_PM_AI_MASTER_WORKBOOK));
        String mf = alt.isEmpty() ? "master.xlsm" : Path.of(alt).getFileName().toString();
        if (mf.startsWith("\\\\")) {
            return primary;
        }
        Path mfPath = Path.of(mf);
        if (mfPath.isAbsolute()) {
            return primary;
        }
        Path pyDir = resolvePythonScriptDir(u);
        Path codeBesidePython = pyDir.getParent();
        Path root = resolveRepoRoot(u);
        Path[] bases =
                new Path[] {
                    codeBesidePython,
                    root.resolve("code"),
                    root.resolve("plan"),
                    root,
                };
        for (Path base : bases) {
            if (base == null || !Files.isDirectory(base)) {
                continue;
            }
            Path c = base.resolve(mf).normalize().toAbsolutePath();
            if (Files.isRegularFile(c)) {
                return c;
            }
        }
        return primary;
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
     * {@code code/} 配下のサマリ用ブックの既定ファイル名（湖南工場プリセット・空欄時の解決にも使用）。実行・ログタブの「開く」から参照。
     */
    public static final String SUMMARY_AI_DISPATCH_XLSX = "サマリ_AI配台.xlsx";

    /**
     * 湖南工場・配台AI {@code 共有DATA} フォルダ（UNC）。{@link FactorySite#KONAN} のマスタ／サマリ既定の親。
     */
    public static final String DEFAULT_KONAN_SHARED_DATA_DIR =
            "\\\\192.168.0.101\\"
                    + "共有フォルダ\\"
                    + "湖南工場\\"
                    + "湖南共有\\"
                    + "002  加工G\\"
                    + "●配台AIシステム\\"
                    + "共有DATA";

    /** {@link FactorySite#KONAN} の {@link #KEY_PM_AI_MASTER_WORKBOOK} 既定（UNC）。 */
    public static final String DEFAULT_PM_AI_MASTER_WORKBOOK_KONAN =
            DEFAULT_KONAN_SHARED_DATA_DIR + "\\master.xlsm";

    /** {@link FactorySite#KONAN} の {@link #KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK} 既定（UNC）。 */
    public static final String DEFAULT_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK_KONAN =
            DEFAULT_KONAN_SHARED_DATA_DIR + "\\" + SUMMARY_AI_DISPATCH_XLSX;

    /**
     * 国分工場・配台AIシステム共有フォルダ（UNC）。{@link #DEFAULT_KOKUBU_DATA_DIR} の親。
     */
    public static final String DEFAULT_KOKUBU_SHARED_DATA_DIR =
            "\\\\192.168.0.101\\"
                    + "共有フォルダ\\"
                    + "国分工場\\"
                    + "国分共有\\"
                    + "●配台AIシステム";

    /** 国分工場 {@code DATA} フォルダ（UNC）。{@link FactorySite#KOKUBU} のマスタ／サマリ既定の親。 */
    public static final String DEFAULT_KOKUBU_DATA_DIR = DEFAULT_KOKUBU_SHARED_DATA_DIR + "\\DATA";

    /** {@link FactorySite#KOKUBU} の {@link #KEY_PM_AI_MASTER_WORKBOOK} 既定（UNC）。 */
    public static final String DEFAULT_PM_AI_MASTER_WORKBOOK_KOKUBU =
            DEFAULT_KOKUBU_DATA_DIR + "\\国分master.xlsm";

    /** {@link FactorySite#KOKUBU} の {@link #KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK} 既定（UNC）。 */
    public static final String DEFAULT_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK_KOKUBU =
            DEFAULT_KOKUBU_DATA_DIR + "\\" + SUMMARY_AI_DISPATCH_XLSX;

    /** {@link FactorySite#KONAN} の {@link #KEY_PM_AI_ALADDIN_MASTER_DIR} 既定（UNC）。 */
    public static final String DEFAULT_PM_AI_ALADDIN_MASTER_DIR_KONAN =
            DEFAULT_KONAN_SHARED_DATA_DIR + "\\" + ALADDIN_MASTER_DIR_LEAF_NAME;

    /** {@link FactorySite#KOKUBU} の {@link #KEY_PM_AI_ALADDIN_MASTER_DIR} 既定（UNC）。 */
    public static final String DEFAULT_PM_AI_ALADDIN_MASTER_DIR_KOKUBU =
            DEFAULT_KOKUBU_DATA_DIR + "\\" + ALADDIN_MASTER_DIR_LEAF_NAME;

    /** {@link FactorySite#KONAN} の {@link #KEY_PM_AI_REQUEST_FORM_JUCHU_FILE} 既定（UNC）。 */
    public static final String DEFAULT_PM_AI_REQUEST_FORM_JUCHU_FILE_KONAN =
            "\\\\192.168.0.101\\"
                    + "共有フォルダ\\"
                    + "湖南工場\\"
                    + "湖南共有\\"
                    + "生産管理システム\\"
                    + "管理システム\\"
                    + DEFAULT_REQUEST_FORM_JUCHU_FILE_NAME;

    /** {@link FactorySite#KOKUBU} の {@link #KEY_PM_AI_REQUEST_FORM_JUCHU_FILE} 既定（UNC）。 */
    public static final String DEFAULT_PM_AI_REQUEST_FORM_JUCHU_FILE_KOKUBU =
            "\\\\192.168.0.101\\"
                    + "共有フォルダ\\"
                    + "国分工場\\"
                    + "国分共有\\"
                    + "加工管理\\"
                    + "加工計画関連\\"
                    + "加工依頼書入力　国分コピー入力（2024年12月28））.xlsm";

    /**
     * @deprecated 国分プリセットのサマリ既定は {@link #DEFAULT_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK_KOKUBU}
     *     （{@link #SUMMARY_AI_DISPATCH_XLSX}）を使用。
     */
    @Deprecated
    public static final String KOKUBU_SUMMARY_AI_DISPATCH_WORKBOOK_XLSX = "国分サマリ_AI配台.xlsx";

    /**
     * リポジトリ {@code code/} 内の {@link #SUMMARY_AI_DISPATCH_XLSX} の絶対パス（{@link #resolveRepoRoot} と同一のルート解決）。
     * {@link #KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK} が非空のときはそのパス（絶対、または {@code code/} 基準の相対）を返す。
     */
    public static Path summaryAiDispatchXlsxPath(Map<String, String> ui) {
        return summaryAiDispatchXlsxPathForFactory(ui, null);
    }

    /**
     * 利用工場に合わせたサマリ Excel パス。
     *
     * <p>環境変数のサマリパスが別工場を指すときは {@code site} の工場既定 UNC を使う（操作者 bin／PDF と整合）。
     */
    public static Path summaryAiDispatchXlsxPathForFactory(Map<String, String> ui, FactorySite site) {
        Map<String, String> u = ui != null ? ui : Map.of();
        if (site != null) {
            String override = trim(u.get(KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK));
            if (!override.isEmpty()) {
                Optional<FactorySite> summarySite =
                        FactorySite.inferFromPortableBundleSourceValue(override);
                if (summarySite.isEmpty() || summarySite.get() == site) {
                    return summaryAiDispatchXlsxPathFromOverride(u, override);
                }
            }
            String factoryDefault = site.pmAiSummaryAiDispatchWorkbookEnvValue(u);
            if (factoryDefault != null && !factoryDefault.isBlank()) {
                return Path.of(factoryDefault.trim()).toAbsolutePath().normalize();
            }
        }
        String override = trim(u.get(KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK));
        if (!override.isEmpty()) {
            return summaryAiDispatchXlsxPathFromOverride(u, override);
        }
        return resolveRepoRoot(u)
                .resolve("code")
                .resolve(SUMMARY_AI_DISPATCH_XLSX)
                .toAbsolutePath()
                .normalize();
    }

    private static Path summaryAiDispatchXlsxPathFromOverride(Map<String, String> u, String override) {
        Path p = Path.of(override);
        if (!p.isAbsolute()) {
            p = resolveRepoRoot(u).resolve("code").resolve(override);
        }
        return p.toAbsolutePath().normalize();
    }

    /** RDP ランチャー／{@link #RDP_LAUNCHER_INI_BASENAME} の配備先（サマリ Excel と同階層）。 */
    public static Path resolveRdpLauncherDeployDir(Map<String, String> ui) {
        Path summary = summaryAiDispatchXlsxPath(ui);
        Path parent = summary.getParent();
        if (parent == null) {
            return summary.toAbsolutePath().normalize();
        }
        return parent.toAbsolutePath().normalize();
    }

    public static Path resolveRdpLauncherExe(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String override = trim(u.get(KEY_PM_AI_RDP_LAUNCHER_EXE));
        if (!override.isEmpty()) {
            return Path.of(override).toAbsolutePath().normalize();
        }
        return resolveRdpLauncherDeployDir(u).resolve(RDP_LAUNCHER_EXE_BASENAME);
    }

    public static Path resolveRdpLauncherIni(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String override = trim(u.get(KEY_PM_AI_RDP_LAUNCHER_INI));
        if (!override.isEmpty()) {
            return Path.of(override).toAbsolutePath().normalize();
        }
        return resolveRdpLauncherDeployDir(u).resolve(RDP_LAUNCHER_INI_BASENAME);
    }

    public static Path resolveRdpLauncherVersionFile(Map<String, String> ui) {
        return resolveRdpLauncherDeployDir(ui).resolve(RDP_LAUNCHER_VERSION_BASENAME);
    }

    /** 学習アーカイブのサブフォルダ名（親は {@link #summaryAiDispatchXlsxPath} と同一）。 */
    public static final String KEY_PM_AI_DISPATCH_LEARNING_ARCHIVE_SUBDIR =
            "PM_AI_DISPATCH_LEARNING_ARCHIVE_SUBDIR";

    public static final String DEFAULT_PM_AI_DISPATCH_LEARNING_ARCHIVE_SUBDIR =
            "dispatch-learning-archive";

    /** 学習アーカイブの背景実行を有効化（将来の学習パイプライン用。現状は手動/archive 更新経路）。 */
    public static final String KEY_PM_AI_LEARNING_ARCHIVE_ENABLED =
            "PM_AI_LEARNING_ARCHIVE_ENABLED";

    /** 実績由来学習速度を配台計画に適用。 */
    public static final String KEY_PM_AI_LEARNED_SPEED_ENABLED = "PM_AI_LEARNED_SPEED_ENABLED";

    /** 学習速度適用に必要な (工程,機械) 別最小観測数。 */
    public static final String KEY_PM_AI_LEARNED_SPEED_MIN_SAMPLES =
            "PM_AI_LEARNED_SPEED_MIN_SAMPLES";

    /** 学習速度のパーセンタイル（既定 50）。 */
    public static final String KEY_PM_AI_LEARNED_SPEED_PERCENTILE = "PM_AI_LEARNED_SPEED_PERCENTILE";

    /** 速度ヒストグラムのビン幅（m/分）。 */
    public static final String KEY_PM_AI_LEARNED_SPEED_HISTOGRAM_BIN_WIDTH =
            "PM_AI_LEARNED_SPEED_HISTOGRAM_BIN_WIDTH";

    /**
     * 学習データ蓄積ルート: {@link #summaryAiDispatchXlsxPath(Map)} の親 +
     * {@link #DEFAULT_PM_AI_DISPATCH_LEARNING_ARCHIVE_SUBDIR}（サブフォルダ名は環境変数で上書き可）。
     */
    public static Path resolveDispatchLearningArchiveRoot(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path parent = summaryAiDispatchXlsxPath(u).getParent();
        if (parent == null) {
            parent = resolveRepoRoot(u).resolve("code");
        }
        String sub = trim(u.get(KEY_PM_AI_DISPATCH_LEARNING_ARCHIVE_SUBDIR));
        if (sub.isEmpty()) {
            sub = DEFAULT_PM_AI_DISPATCH_LEARNING_ARCHIVE_SUBDIR;
        }
        return parent.resolve(sub).toAbsolutePath().normalize();
    }

    /** 設備ガント PDF（{@link #summaryAiDispatchXlsxPath} と同一フォルダ）。VBA スナップショット名に合わせる。 */
    public static final String EQUIPMENT_GANTT_PDF = "結果_設備ガント.pdf";

    /** 実行時間分析タブの履歴 JSON（{@link #summaryAiDispatchXlsxPath} と同一フォルダ）。 */
    public static final String PIPELINE_EXECUTION_TIMING_HISTORY_JSON =
            "pipeline-execution-timing-history.json";

    /**
     * 工場別操作者名・PIN 設定（{@link FactoryOperatorUserStore}）。{@link #summaryAiDispatchXlsxPath(Map)}
     * と同一フォルダのバイナリ。
     */
    public static final String FACTORY_OPERATOR_USERS_BIN = "factory-operator-users.bin";

    /**
     * サマリ Excel 世代退避フォルダ（{@link #summaryAiDispatchXlsxPath(Map)} の親配下）。
     * {@link jp.co.pm.ai.desktop.io.SummaryAiDispatchGenerationStore} が使用。
     */
    public static final String SUMMARY_AI_DISPATCH_GENERATIONS_DIR = "summary-ai-dispatch-generations";

    /**
     * 実行時間履歴 JSON の絶対パス。親フォルダは {@link #summaryAiDispatchXlsxPath(Map)} と同一。
     */
    public static Path pipelineExecutionTimingHistoryPath(Map<String, String> ui) {
        return siblingOfSummaryAiDispatchWorkbook(ui, PIPELINE_EXECUTION_TIMING_HISTORY_JSON);
    }

    /**
     * ユーザー管理バイナリの手動バックアップ退避フォルダ（{@link #factoryOperatorUsersStorePath(Map)} の親配下）。
     */
    public static final String FACTORY_OPERATOR_USERS_BACKUPS_DIR = "factory-operator-users-backups";

    /**
     * 依頼書入力・受注ファイル Excel のローカル世代バックアップ（{@link
     * jp.co.pm.ai.desktop.io.RequestFormJuchuFileBackupStore}）。
     */
    public static final String REQUEST_FORM_JUCHU_FILE_BACKUPS_DIR = "request-form-juchu-backups";

    /**
     * 操作者名・PIN 設定の絶対パス。親フォルダは {@link #summaryAiDispatchXlsxPath(Map)} と同一。
     */
    public static Path factoryOperatorUsersStorePath(Map<String, String> ui) {
        return factoryOperatorUsersStorePath(ui, null);
    }

    /**
     * 利用工場に合わせた操作者名・PIN 設定の絶対パス。
     *
     * <p>環境変数のサマリパスが別工場を指すときは {@code site} の工場既定 UNC 配下を使う。
     */
    public static Path factoryOperatorUsersStorePath(Map<String, String> ui, FactorySite site) {
        return siblingOfSummaryAiDispatchWorkbookForFactory(ui, site, FACTORY_OPERATOR_USERS_BIN);
    }

    /**
     * 操作者名・PIN のローカル退避（{@code ~/.pm-ai-desktop/factory-operator-users-<site>.bin}）。
     *
     * <p>工場別 UNC の共有 DATA に書き込み権限が無い／別工場パスを参照しているときのフォールバック。
     * 複数 PC で一覧を共有する正本は引き続き {@link #factoryOperatorUsersStorePath(Map, FactorySite)}（ネットワーク側）。
     */
    public static Path localFactoryOperatorUsersStorePath(FactorySite site) {
        FactorySite effective = site != null ? site : FactorySite.KONAN;
        String suffix = effective.name().toLowerCase(Locale.ROOT);
        return Paths.get(
                        System.getProperty("user.home"),
                        ".pm-ai-desktop",
                        "factory-operator-users-" + suffix + ".bin")
                .toAbsolutePath()
                .normalize();
    }

    /**
     * 操作者ダイアログで最後に選んだメンバー名（PC ローカル。{@link #factoryOperatorUsersStorePath} とは別）。
     * 次回起動時に同一 PC で操作者ダイアログを省略して復元するために使う。
     */
    public static Path localFactoryOperatorLastSelectedPath(FactorySite site) {
        FactorySite effective = site != null ? site : FactorySite.KONAN;
        String suffix = effective.name().toLowerCase(Locale.ROOT);
        return Paths.get(
                        System.getProperty("user.home"),
                        ".pm-ai-desktop",
                        "last-factory-operator-" + suffix + ".txt")
                .toAbsolutePath()
                .normalize();
    }

    /** 工場別ユーザー管理 PDF のファイル名（{@link FactorySite#name()} を含む）。 */
    public static String factoryOperatorUsersPdfFileName(FactorySite site) {
        FactorySite effective = site != null ? site : FactorySite.KONAN;
        return "factory-operator-users-" + effective.name() + ".pdf";
    }

    /**
     * 工場別ユーザー管理 PDF の絶対パス。親フォルダは {@link #summaryAiDispatchXlsxPath(Map)} と同一。
     */
    public static Path factoryOperatorUsersPdfPath(Map<String, String> ui, FactorySite site) {
        FactorySite effective = site != null ? site : FactorySite.KONAN;
        return siblingOfSummaryAiDispatchWorkbookForFactory(
                ui, effective, factoryOperatorUsersPdfFileName(effective));
    }

    /**
     * ユーザー管理バイナリの手動バックアップルート。{@link #factoryOperatorUsersStorePath(Map)} の親配下。
     */
    public static Path factoryOperatorUsersBackupsRoot(Map<String, String> ui) {
        return factoryOperatorUsersBackupsRoot(ui, null);
    }

    /**
     * 依頼書入力・受注ファイル Excel のローカル世代バックアップルート（{@code ~/.pm-ai-desktop/…}）。
     *
     * <p>ネットワーク上の受注ファイルを書き込む前に、同一 PC のローカルへ退避する。
     */
    public static Path requestFormJuchuFileBackupsRoot(Map<String, String> ui) {
        FactorySite site = GlobalInitSettingTarget.loadEffective(ui != null ? ui : Map.of());
        String suffix = site.name().toLowerCase(Locale.ROOT);
        return Paths.get(
                        System.getProperty("user.home"),
                        ".pm-ai-desktop",
                        REQUEST_FORM_JUCHU_FILE_BACKUPS_DIR,
                        suffix)
                .toAbsolutePath()
                .normalize();
    }

    /** 利用工場に合わせた手動バックアップルート。 */
    public static Path factoryOperatorUsersBackupsRoot(Map<String, String> ui, FactorySite site) {
        Path store = factoryOperatorUsersStorePath(ui, site);
        Path parent = store.getParent();
        if (parent == null) {
            return resolveRepoRoot(ui)
                    .resolve("code")
                    .resolve(FACTORY_OPERATOR_USERS_BACKUPS_DIR)
                    .toAbsolutePath()
                    .normalize();
        }
        return parent.resolve(FACTORY_OPERATOR_USERS_BACKUPS_DIR).toAbsolutePath().normalize();
    }

    /**
     * サマリ Excel 世代退避のルート。{@link #summaryAiDispatchXlsxPath(Map)} の親配下。
     */
    public static Path summaryAiDispatchGenerationsRoot(Map<String, String> ui) {
        Path summary = summaryAiDispatchXlsxPath(ui);
        Path parent = summary.getParent();
        if (parent == null) {
            return resolveRepoRoot(ui)
                    .resolve("code")
                    .resolve(SUMMARY_AI_DISPATCH_GENERATIONS_DIR)
                    .toAbsolutePath()
                    .normalize();
        }
        return parent.resolve(SUMMARY_AI_DISPATCH_GENERATIONS_DIR).toAbsolutePath().normalize();
    }

    /**
     * 設備ガント PDF の絶対パス。親フォルダは {@link #summaryAiDispatchXlsxPath(Map)} と同一。
     */
    public static Path equipmentGanttPdfPath(Map<String, String> ui) {
        return siblingOfSummaryAiDispatchWorkbook(ui, EQUIPMENT_GANTT_PDF);
    }

    private static Path siblingOfSummaryAiDispatchWorkbook(Map<String, String> ui, String fileName) {
        return siblingOfSummaryAiDispatchWorkbookForFactory(ui, null, fileName);
    }

    private static Path siblingOfSummaryAiDispatchWorkbookForFactory(
            Map<String, String> ui, FactorySite site, String fileName) {
        Path summary = summaryAiDispatchXlsxPathForFactory(ui, site);
        Path parent = summary.getParent();
        if (parent == null) {
            return resolveRepoRoot(ui).resolve("code").resolve(fileName).toAbsolutePath().normalize();
        }
        return parent.resolve(fileName).toAbsolutePath().normalize();
    }

    /** @deprecated {@link #summaryAiDispatchXlsxPath(Map)} を使用 */
    @Deprecated
    public static Path summaryAiDispatchXlsmPath(Map<String, String> ui) {
        return summaryAiDispatchXlsxPath(ui);
    }

    /** @deprecated {@link #SUMMARY_AI_DISPATCH_XLSX} を使用 */
    @Deprecated
    public static final String SUMMARY_AI_DISPATCH_XLSM = SUMMARY_AI_DISPATCH_XLSX;

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
     * Written by {@code run_stage1_extract} beside {@link #summaryAiDispatchXlsxPath} ({@code
     * STAGE1_EXCLUDE_RULES_JSON_FILENAME}).
     */
    public static final String STAGE1_EXCLUDE_RULES_JSON_FILENAME = "stage1_exclude_rules.json";

    /**
     * §B 特別ルール JSON 作業正本（サマリ Excel 同階層 {@code dispatch_special_rules/}）。
     */
    public static Path dispatchSpecialRulesJsonPath(Map<String, String> ui) {
        return jp.co.pm.ai.desktop.dispatch.rules.paths.DispatchRulePaths.workJsonPath(ui);
    }

    /** 作業先が無ければリポジトリ同梱テンプレからコピー。 */
    public static boolean ensureDispatchSpecialRulesJsonFromRepoIfMissing(Map<String, String> ui) {
        return jp.co.pm.ai.desktop.dispatch.rules.paths.DispatchRulePaths
                .ensureWorkJsonFromRepoIfMissing(ui);
    }

    public static java.util.Optional<Path> resolveDefaultDispatchSpecialRulesJsonPath(
            Map<String, String> ui) {
        return jp.co.pm.ai.desktop.dispatch.rules.paths.DispatchRulePaths
                .resolveDefaultWorkJson(ui);
    }

    /**
     * Gemini List Models から抽出した Flash-Lite 系（無料枠向け）のキャッシュ JSON（{@code code/json/} 配下）。
     */
    public static final String GEMINI_FREE_TIER_FLASH_LITE_MODELS_JSON_FILENAME =
            "gemini_free_tier_flash_lite_models.json";

    /**
     * {@link #GEMINI_FREE_TIER_FLASH_LITE_MODELS_JSON_FILENAME} の絶対パス（日次バックグラウンド更新の永続先）。
     */
    public static Path geminiFreeTierFlashLiteModelsCachePath(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        return resolveRepoRoot(u)
                .resolve("code")
                .resolve("json")
                .resolve(GEMINI_FREE_TIER_FLASH_LITE_MODELS_JSON_FILENAME)
                .toAbsolutePath()
                .normalize();
    }

    /**
     * Path: {@link #summaryAiDispatchXlsxPath(Map)} と同一フォルダの段階1配台不要ルール JSON。
     */
    public static Path stage1ExcludeRulesJsonPath(Map<String, String> ui) {
        return siblingOfSummaryAiDispatchWorkbook(ui, STAGE1_EXCLUDE_RULES_JSON_FILENAME);
    }

    /**
     * 依頼書入力タブの ComboBox 候補・受注ファイルパス等（{@link
     * jp.co.pm.ai.desktop.reconciliation.RequestFormInputSettingsStore}）。
     */
    public static final String REQUEST_FORM_INPUT_SETTINGS_JSON_FILENAME =
            "request_form_input_settings.json";

    /** {@link #summaryAiDispatchXlsxPath(Map)} と同一フォルダの依頼書入力設定 JSON。 */
    public static Path requestFormInputSettingsJsonPath(Map<String, String> ui) {
        return siblingOfSummaryAiDispatchWorkbook(ui, REQUEST_FORM_INPUT_SETTINGS_JSON_FILENAME);
    }

    /**
     * リポジトリ同梱の {@code code/json/stage1_exclude_rules.json}（作業コピー元・読込フォールバック）。
     */
    public static Path stage1ExcludeRulesJsonPathLegacyUnderCodeJson(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        return resolveRepoRoot(u)
                .resolve("code")
                .resolve("json")
                .resolve(STAGE1_EXCLUDE_RULES_JSON_FILENAME)
                .toAbsolutePath()
                .normalize();
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
     * リポジトリ内の配台不要ルール JSON テンプレート（{@code code/exclude_rules.json} を優先、無ければ
     * {@code code/json/stage1_exclude_rules.json}）。
     */
    public static Optional<Path> resolveBundledExcludeRulesJsonSourceInRepo(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path primary = resolveRepoRoot(u).resolve("code").resolve("exclude_rules.json");
        if (Files.isRegularFile(primary)) {
            return Optional.of(primary.toAbsolutePath().normalize());
        }
        Path stage1 = stage1ExcludeRulesJsonPathLegacyUnderCodeJson(u);
        if (Files.isRegularFile(stage1)) {
            return Optional.of(stage1.toAbsolutePath().normalize());
        }
        return Optional.empty();
    }

    /**
     * 作業先（サマリ Excel と同一フォルダ）に {@link #STAGE1_EXCLUDE_RULES_JSON_FILENAME} が無ければ、
     * リポジトリ同梱または旧配置からコピーする。
     *
     * @return 作業先ファイルが実在するとき {@code true}
     */
    public static boolean ensureStage1ExcludeRulesJsonFromRepoIfMissing(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path target = stage1ExcludeRulesJsonPath(u);
        if (Files.isRegularFile(target)) {
            return true;
        }
        Optional<Path> source = resolveBundledExcludeRulesJsonSourceInRepo(u);
        if (source.isEmpty()) {
            Path legacyCodeJson = stage1ExcludeRulesJsonPathLegacyUnderCodeJson(u);
            if (Files.isRegularFile(legacyCodeJson) && !legacyCodeJson.equals(target)) {
                source = Optional.of(legacyCodeJson);
            } else {
                Path legacyPythonJson = stage1ExcludeRulesJsonPathLegacyUnderPython(u);
                if (Files.isRegularFile(legacyPythonJson) && !legacyPythonJson.equals(target)) {
                    source = Optional.of(legacyPythonJson);
                }
            }
        }
        if (source.isEmpty()) {
            return false;
        }
        try {
            if (target.getParent() != null) {
                Files.createDirectories(target.getParent());
            }
            Files.copy(source.get(), target, StandardCopyOption.REPLACE_EXISTING);
            return Files.isRegularFile(target);
        } catch (IOException ex) {
            return false;
        }
    }

    /**
     * Default for {@link #KEY_PM_AI_EXCLUDE_RULES_JSON}: {@link #summaryAiDispatchXlsxPath(Map)} と同一フォルダの
     * {@link #STAGE1_EXCLUDE_RULES_JSON_FILENAME}。無ければ {@link #ensureStage1ExcludeRulesJsonFromRepoIfMissing} で
     * リポジトリ同梱からコピーしてから返す。
     */
    public static Optional<Path> resolveDefaultExcludeRulesJsonPath(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        if (ensureStage1ExcludeRulesJsonFromRepoIfMissing(u)) {
            return Optional.of(stage1ExcludeRulesJsonPath(u));
        }
        return Optional.empty();
    }

    /** 材料・製品種類ルックアップ表（サマリ Excel と同一フォルダに配置）。 */
    public static final String DISPATCH_LOOKUP_USED_RAW_ROLL = "使用原反,ロール単位の長さ.txt";

    public static final String DISPATCH_LOOKUP_PRODUCT_ROLL = "製品名,ロール単位の長さ.txt";
    public static final String DISPATCH_LOOKUP_PRODUCT_WIDTH = "製品名, 製品幅.txt";
    public static final String DISPATCH_LOOKUP_PRODUCT_THICK = "製品名,製品厚み.txt";
    public static final String DISPATCH_LOOKUP_PRODUCT_LENGTH = "製品名,製品長.txt";
    public static final String DISPATCH_LOOKUP_USED_RAW_WIDTH = "使用原反, 加工幅.txt";

    private static final List<String> DISPATCH_LOOKUP_TABLE_FILENAMES =
            List.of(
                    DISPATCH_LOOKUP_USED_RAW_ROLL,
                    DISPATCH_LOOKUP_PRODUCT_ROLL,
                    DISPATCH_LOOKUP_PRODUCT_WIDTH,
                    DISPATCH_LOOKUP_PRODUCT_THICK,
                    DISPATCH_LOOKUP_PRODUCT_LENGTH,
                    DISPATCH_LOOKUP_USED_RAW_WIDTH);

    /**
     * 材料・製品種類ルックアップ表の絶対パス（{@link #summaryAiDispatchXlsxPath(Map)} と同一フォルダ）。
     */
    public static Path dispatchLookupTablePath(Map<String, String> ui, String filename) {
        return siblingOfSummaryAiDispatchWorkbook(ui, filename);
    }

    /** リポジトリ {@code code/} 配下の同梱テーブル（コピー元）。 */
    public static Path dispatchLookupTablePathLegacyUnderCode(Map<String, String> ui, String filename) {
        return resolveCodeDir(ui != null ? ui : Map.of())
                .resolve(filename)
                .toAbsolutePath()
                .normalize();
    }

    public static Optional<Path> resolveBundledDispatchLookupTableSourceInRepo(
            Map<String, String> ui, String filename) {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path underCode = dispatchLookupTablePathLegacyUnderCode(u, filename);
        if (Files.isRegularFile(underCode)) {
            return Optional.of(underCode);
        }
        Path atRepo = resolveRepoRoot(u).resolve("code").resolve(filename);
        if (Files.isRegularFile(atRepo)) {
            return Optional.of(atRepo.toAbsolutePath().normalize());
        }
        return Optional.empty();
    }

    /**
     * 作業先（サマリ Excel 同フォルダ）にテーブルが無ければリポジトリ {@code code/} からコピーする。
     *
     * @return 作業先ファイルが実在するとき {@code true}
     */
    public static boolean ensureDispatchLookupTableFromRepoIfMissing(
            Map<String, String> ui, String filename) {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path target = dispatchLookupTablePath(u, filename);
        if (Files.isRegularFile(target)) {
            return true;
        }
        Optional<Path> source = resolveBundledDispatchLookupTableSourceInRepo(u, filename);
        if (source.isEmpty()) {
            Path legacy = dispatchLookupTablePathLegacyUnderCode(u, filename);
            if (Files.isRegularFile(legacy) && !legacy.equals(target)) {
                source = Optional.of(legacy);
            }
        }
        if (source.isEmpty()) {
            return false;
        }
        try {
            if (target.getParent() != null) {
                Files.createDirectories(target.getParent());
            }
            Files.copy(source.get(), target, StandardCopyOption.REPLACE_EXISTING);
            return Files.isRegularFile(target);
        } catch (IOException ex) {
            return false;
        }
    }

    /** {@link #DISPATCH_LOOKUP_TABLE_FILENAMES} をすべて作業先へ確保する。 */
    public static void ensureAllDispatchLookupTablesFromRepoIfMissing(Map<String, String> ui) {
        for (String filename : DISPATCH_LOOKUP_TABLE_FILENAMES) {
            ensureDispatchLookupTableFromRepoIfMissing(ui, filename);
        }
    }

    /** 材料・製品種類ルックアップ表のファイル名一覧（作業先・リポジトリ同梱の対象）。 */
    public static List<String> dispatchLookupTableFilenames() {
        return DISPATCH_LOOKUP_TABLE_FILENAMES;
    }

    /** リポジトリ同梱から作業先へ上書きコピーした結果。 */
    public record DispatchLookupTableOverwriteResult(
            String filename,
            boolean success,
            String message,
            Path sourcePath,
            Path targetPath) {}

    /**
     * リポジトリ {@code code/} 同梱のルックアップ表で作業先（サマリ Excel 同フォルダ）を上書きする。
     *
     * <p>作業先が既にあっても {@link StandardCopyOption#REPLACE_EXISTING} で置き換える。
     */
    public static DispatchLookupTableOverwriteResult overwriteDispatchLookupTableFromRepo(
            Map<String, String> ui, String filename) {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path target = dispatchLookupTablePath(u, filename);
        Optional<Path> source = resolveBundledDispatchLookupTableSourceInRepo(u, filename);
        if (source.isEmpty()) {
            return new DispatchLookupTableOverwriteResult(
                    filename,
                    false,
                    "リポジトリに同梱ファイルがありません",
                    null,
                    target);
        }
        try {
            if (target.getParent() != null) {
                Files.createDirectories(target.getParent());
            }
            Files.copy(source.get(), target, StandardCopyOption.REPLACE_EXISTING);
            return new DispatchLookupTableOverwriteResult(
                    filename, true, "上書きしました", source.get(), target);
        } catch (IOException ex) {
            return new DispatchLookupTableOverwriteResult(
                    filename, false, ex.getMessage(), source.get(), target);
        }
    }

    /** {@link #dispatchLookupTableFilenames()} をすべて作業先へ上書きコピーする。 */
    public static List<DispatchLookupTableOverwriteResult> overwriteAllDispatchLookupTablesFromRepo(
            Map<String, String> ui) {
        List<DispatchLookupTableOverwriteResult> out = new ArrayList<>();
        for (String filename : DISPATCH_LOOKUP_TABLE_FILENAMES) {
            out.add(overwriteDispatchLookupTableFromRepo(ui, filename));
        }
        return out;
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

    /** {@link #resolveRepoRoot(Map)}/{@link #MANUAL_INDEX_HTML_REL}（ブラウザで開く取扱説明書トップ）。 */
    public static Path resolveManualIndexHtml(Map<String, String> ui) {
        return resolveRepoRoot(ui).resolve(MANUAL_INDEX_HTML_REL).toAbsolutePath().normalize();
    }

    /** {@link #resolveRepoRoot(Map)}/{@link #DISPATCH_USAGE_GUIDE_DOCX}（Word で開く現場手順書）。 */
    public static Path resolveDispatchUsageGuideDocx(Map<String, String> ui) {
        return resolveRepoRoot(ui).resolve(DISPATCH_USAGE_GUIDE_DOCX).toAbsolutePath().normalize();
    }

    /** {@link #resolveRepoRoot(Map)}/{@link #DISPATCH_RULES_HTML_REL}（ブラウザで開く配台ルール）。 */
    public static Path resolveDispatchRulesHtml(Map<String, String> ui) {
        return resolveRepoRoot(ui).resolve(DISPATCH_RULES_HTML_REL).toAbsolutePath().normalize();
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
