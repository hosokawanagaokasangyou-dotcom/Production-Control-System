package jp.co.pm.ai.desktop.config;

import java.util.HashMap;
import java.util.Map;

/**
 * Supplemental descriptions derived from {@code workbook_env_bootstrap.py}, {@code planning_core}, and the
 * desktop bridge (not from OS env). Merged with sheet text in the UI.
 *
 * <p>Variables whose names mention xlwings concern Excel COM automation from the Python stack when Excel
 * invokes those scripts (add-in / legacy macro workflows). They are not prerequisites for the JavaFX
 * desktop launcher path, which runs child Python headlessly for stages 1/2 without xlwings.
 */
public final class EnvVarDocs {

    private static final Map<String, String> LOGIC = new HashMap<>();

    static {
        put(
                "PM_AI_PYTHON",
                "段階1/2 等の子プロセスで使う Python 実行ファイル（パスは実行ファイル。フォルダのみ指定すると実行時に python.exe 等へ補正）。"
                        + "編集は環境変数タブのみ（実行・ログタブに Python 入力は無い）。"
                        + "値が空のとき: まず user.dir 周辺で pm-ai-data/runtime/python-embed/python.exe が実在すればそれを使い、"
                        + "無ければ OS の PATH 上の python / python3（開発ツリーに python-embed を置いていない場合はここに落ちる）。"
                        + "複数 Python を PATH で切り替えている／特定の python.exe に固定したいときは、本変数にその絶対パスを明示する。"
                        + "環境変数タブを空にしても動くのはこのフォールバックのため。"
                        + "初期化・空欄補完では pm-ai-data/runtime/python-embed/python.exe を user.dir から親ディレクトリへ最大12段まで辿って探索し、見つかれば絶対パスで入れる（ショートカットで user.dir が bin 等になる場合のため）。"
                        + "見つからずインストール根がポータル配布なら相対パス。開発などでは同梱 exe が取れたら絶対パス、無ければ PATH の python/python3。");
        put(
                "PM_AI_CODE_PYTHON_DIR",
                "スクリプト根（task_extract_stage1.py 等）。"
                        + "自動検出は user.dir から code/python を探す。");
        put(
                "PM_AI_REPO_ROOT",
                "Production-Control-System の親（リポジトリ根）。"
                        + "PM_AI_CODE_PYTHON_DIR 未指定時の推定に使用。");
        put(
                "PM_AI_PORTABLE_BUNDLE_SOURCE_DIR",
                "ポータブル配布（PMD.exe と pm-ai-data）向け。正本は次のいずれか。"
                        + "（1）リポジトリルートのフォルダパス（UNC 可）。直下の version.txt とローカル pm-ai-data を比較し、新しいときのみ起動時に pm-ai-data を同期する。"
                        + "（2）バージョンアップ用の .zip ファイルのパス（配布は固定名 PMD_version_upgrade.zip を推奨。バージョンは ZIP と同じフォルダの外付け version.txt で区別）。起動時に ZIP を自動展開してから pm-ai-data に同期する。"
                        + "空のときは自動更新しない（情報表示のみ）。");
        put(
                "PM_AI_OUTPUT_DIR",
                "段階1/2 の出力先（plan_input_tasks.xlsx 等、従来 code/output"
                        + " に相当）。未設定時は PM_AI_REPO_ROOT"
                        + " 直下の output（JavaFX と planning_core.bootstrap と同解決）。");
        put(
                "PM_AI_WORKSPACE",
                "配台作業ルート（Python の cwd、ログ/output、"
                        + "Gemini 証明書の搜索先。JavaFX と planning_core.bootstrap "
                        + "で最優先される。"
                        + "未指定時は PM_AI_CODE_PYTHON_DIR の親（code）から"
                        + "推定する場合が多い。");
        put(
                "PM_AI_PROCESSING_PLAN_PATH",
                "段階1用：加工計画DATA相当の表（CSV/Parquet/xlsx）。"
                        + "Python は未設定またはファイル無しのとき"
                        + "、PM_AI_TASK_INPUT_SOURCE_DIR 内の最新表を自動で"
                        + "この変数に設定（dispatch_workspace.resolve_processing_plan_path_from_env）。"
                        + "run_stage1_extract はこのパス（または SOURCE_DIR"
                        + "解決の実在ファイル）が必要。"
                        + "配台不要は master.xlsm から json/stage1_exclude_rules.json に書き出し。"
                        + " 正式な列構成は plan/01_加工計画DATA_単一ファイル.m"
                        + " と同等の Power Query 成形後の加工計画DATA相当。"
                        + "生の問合せ xlsx を直接指定する場合は Python"
                        + " 側でヘッダー行・列名の救済のみ（PQ"
                        + " の複合見出しや日付列名の展開は再現しない）。"
                        + "確実に合わせるときはクエリ更新後の CSV"
                        + " 等のパスを指定すること。");
        put(
                "PM_AI_PLAN_INPUT_PATH",
                "専用UIで指定した配台計画タスク入力ファイルへのパス"
                        + "（CSV / Parquet / xlsx / xlsm 等）。段階2の"
                        + " load_planning_tasks_df は表形式を読むため"
                        + "、必ずしもExcelブックではない。"
                        + "JavaFX での運用ではxlwings（Excel アドイン連携用の COM操作）は本アプリの必須ではない。"
                        + "Excel から起動するPython経路でブックを開くときのみ"
                        + "、対象処理に実在するxlsx/xlsm が役に立つ。"
                        + "設定時はマクロブックのそのシートを元にしない。");
        put(
                "PM_AI_PROCESSING_PLAN_SHEET",
                "PM_AI_PROCESSING_PLAN_PATH が xlsx のときのシート指定。空で"
                        + "先頭シート（0番）。単一シートなら"
                        + "名前不要。複数シートで名前を指す場合は"
                        + "文字列。数値のみ（例: 1）は 0始まりの"
                        + "インデックス。");
        put(
                "PM_AI_PROCESSING_PLAN_HEADER_ROW",
                "xlsx 読込み時の列名行（Excel の 1 始まりの"
                        + "行番号）。空で、同一行に「依頼NO」"
                        + "と「工程名」ある最上位の行を"
                        + "自動探知（工程別生産計画問合せ"
                        + "など先頭にメタ行があるブックは"
                        + "通常 6 行目）。");
        put(
                "PM_AI_KOUBAI_INQUIRY_SHAPING",
                "工程別問合せ"
                        + " xlsx: "
                        + "6+5"
                        + "行"
                        + "複合見出し"
                        + "、"
                        + "加工時間"
                        + "/"
                        + "加工速度"
                        + "列削除"
                        + "、"
                        + "加工数量"
                        + "の部分除去"
                        + "（見出しが「加工数量」のみの列は列名維持）"
                        + "、"
                        + "YYYY/MM/DD"
                        + "。"
                        + "空"
                        + "=auto, 0=off, 1=force.");
        put(
                "PM_AI_TABULAR_CSV_ENCODING",
                "PM_AI_PROCESSING_PLAN_PATH 等 CSV の文字コード（空で utf-8-sig）。");
        put(
                "PM_AI_GLOBAL_PRIORITY_OVERRIDE_PATH",
                "段階2 メイン「グローバルコメント」代替: UTF-8"
                        + " テキストファイル1本（パスあれば"
                        + " Excel シートスキャンなし）。"
                        + " input_resolution / load_main_sheet_global_priority_override_text。");
        put(
                "PM_AI_RESULT_TASK_COLUMN_CONFIG_CSV",
                "結果_タスク一覧の列設定（列名、表示"
                        + "列を持つ CSV。あれば列設定シート読みをスキップ。");
        put(
                "PM_AI_COLUMN_CONFIG_WORKBOOK",
                "列設定_結果_タスク一覧シートを含む"
                        + " xlsx/xlsm。PM_AI_PLAN_INPUT_PATH と異なる列設定専用ブック"
                        + "を指す場合。");
        put(
                "PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK",
                "加工計画DATA等から「データ抽出時間」列を"
                        + "読むブック（未指定時は planning_core の"
                        + " input_resolution による探索、PM_AI_PLAN_INPUT_PATH など）。");
        put(
                "PM_AI_ACTUALS_DATA_WORKBOOK",
                "加工実績DATA シートを読むブック。"
                        + "未設定時は PM_AI_ACTUAL_DETAIL_WORKBOOK →"
                        + " PM_AI_ACTUAL_DETAIL_SOURCE_DIR 内最新 xlsx/xlsm"
                        + " → PM_AI_PLAN_INPUT_PATH がExcelのときそのブック"
                        + " と実績明細と同じ既定探索（input_resolution）。");
        put(
                "PM_AI_ACTUALS_DATA_SHEET",
                "PM_AI_ACTUALS_DATA_WORKBOOK 内のシート指定。空で"
                        + "先頭シート（0番）。単一シートなら名前不要。"
                        + "数値のみは 0始まりのインデックス。");
        put(
                "PM_AI_ACTUAL_DETAIL_SHEET",
                "PM_AI_ACTUAL_DETAIL_WORKBOOK 等で読む加工実績明細のシート指定。空で"
                        + "先頭シート（0番）。単一シートなら名前不要。"
                        + "数値のみは 0始まりのインデックス。");
        put(
                "PM_AI_TASK_INPUT_SOURCE_DIR",
                "PQ-A 加工計画DATA取得元（plan/01_*.m の Folder.Files と同系）。"
                        + "未設定時は \\\\192.168.0.101\\共有...●DATA\\生産計画問合せ。"
                        + "JavaFX 初期値は AppPaths.resolveTaskInputSourceDir。"
                        + "Python は PM_AI_PROCESSING_PLAN_PATH が未設定または存在しないとき"
                        + "、このフォルダ内 CSV/Parquet/xlsx 等のうち"
                        + "更新時刻が最新の1件をタスク入力に使用。"
                        + "「納期管理ビュー」タブの計画側アラジン数量もこの解決パス由来の加工計画シートを参照する。");
        put(
                "PM_AI_ACTUAL_DETAIL_SOURCE_DIR",
                "加工実績明細DATA 出力元（plan/02__q*.m の Folder.Files と同系）。"
                        + "planning_core はこのフォルダ内の最新 xlsx/xlsm"
                        + " を実績明細読込の元にする。"
                        + " PM_AI_ACTUALS_DATA_WORKBOOK 未設定時は"
                        + "加工実績DATA 読込も同じ最新ファイルを使用。"
                        + "未設定時は 002  加工G\\●検査表作成\\加工実績明細DATA系 UNC。"
                        + "PM_AI_ACTUAL_DETAIL_WORKBOOK で単一ファイルを優先。");
        put(
                "PM_AI_ACTUAL_DETAIL_WORKBOOK",
                "加工実績明細DATAを読むブックのフルパス（指定時は"
                        + " PM_AI_ACTUAL_DETAIL_SOURCE_DIR より優先）。");
        put(
                "PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES",
                "JavaFX「加工実績」明細タブ: 元 Excel/CSV を POI で読む前のファイルサイズ上限（バイト）。"
                        + "超過時は読込を中止しメッセージ表示（OOM 回避）。空または未設定で 20971520（20MiB）。"
                        + "0 以下で上限なし。例: 16777216、16M、64MB。");
        put(
                "PM_AI_RESULT_DISPATCH_TABLE_DIR",
                "Power Query _q結果_配台表 参照用の"
                        + " 結果_配台表.xlsx 出力先（マクロブック側に"
                        + " フォルダパス名を合わせる場合）。"
                        + "未設定時は段階2は PM_AI_WORKSPACE または"
                        + " PM_AI_PLAN_INPUT_PATH 親階層に合わせる場合がある、"
                        + "JavaFX 初期値は PM_AI_REPO_ROOT 下の code/output（例: Production-Control-System/code/output/"
                        + "結果_配台表.xlsx 同階層に 結果_配台表.json も出力）。"
                        + "「納期管理ビュー」の計画比較サブタブはこのフォルダ直下の 結果_配台表.json と"
                        + "タスク入力ソースのアラジン日別数量を突き合わせる。");
        put(
                "GANTT_ACTUAL_DETAIL_DATE_FROM",
                "納期管理ビュー／実績明細ガント共通：実績側で表示する暦日の開始（空＝下限なし）。"
                        + " planning_core の ENV と同じ。");
        put(
                "GANTT_ACTUAL_DETAIL_DATE_TO",
                "納期管理ビュー／実績明細ガント共通：実績側で表示する暦日の終了（空＝上限なし）。");
        put(
                "PM_AI_RESULT_DISPATCH_TABLE_JSON",
                "段階2 の 結果_配台表.json 出力："
                        + "0/false/no/off/none で無効（空で有効、xlsx と同データ）。");
        put(
                "PM_AI_EXCEL_TRACE_TASK_ID",
                "段階2（配台試行含む）の Excel 生成経路を 1 依頼で追跡するデバッグ用依頼NO（例: Y5-14）。"
                        + "本アプリの「環境変数」タブにのみ設定（子プロセスへ引き渡し）。"
                        + "OS の PM_AI_* は起動に使わない（空のまま推奨）。"
                        + "有効時は .cursor/debug-excel-trace.log に NDJSON（EX1=df_tasks、EX4=サイドカー JSON、"
                        + "EX5=両者のセル差分）。全ブック JSON（PM_AI_PLAN_WORKBOOK_JSON）とは別。");
        put(
                "GEMINI_CREDENTIALS_JSON",
                "Gemini 暗号化証明書 JSON（例: gemini_credentials.encrypted.json）の"
                        + "フルパス。planning_core で最優先。"
                        + "JavaFX 環境変数タブの「ファイル...」で選択可。");
        put(
                "PM_AI_MASTER_WORKBOOK",
                "master 系 .xlsm の絶対パス（実在ファイルのとき"
                        + " MASTER_WORKBOOK_FILE より優先。planning_core の"
                        + " マスタ読込・機械カレンダー等に使用。"
                        + " JavaFX の「マスタ読込サマリ」タブで内容を確認可。");
        put(
                "PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK",
                "実行・ログタブの「開く」（サマリ AI 配台"
                        + " 等）が開くブックの絶対パス（.xlsx 等）。"
                        + " 空で code/ 下の"
                        + " サマリ_AI配台.xlsx（"
                        + "PM_AI_REPO_ROOT 準拠）。"
                        + " ファイル名のみのときは code/ からの相対パス"
                        + "として解決。");
        put(
                "PM_AI_SKIP_WORKBOOK_ENV_SHEET",
                "1/true 等で workbook_env_bootstrap がマクロブックの"
                        + "「設定_環境変数」シートを読まない。"
                        + "JavaFX 環境変数タブが子プロセスの源。"
                        + " 空のときランチャーは 1 を付与。"
                        + " OS 環境変数へは書き込まない運用を前提。");
        put(
                "PM_AI_EXCLUDE_RULES_JSON",
                "段階1（run_stage1_extract）で master.xlsm 「設定_配台不要工程」"
                        + "を json/stage1_exclude_rules.json へ書き出し、本変数を"
                        + " その絶対パスに自動設定（子プロセス内）。"
                        + " 手動でも UTF-8 JSON（list または {\"rules\":[...]}、"
                        + " 列構造は設定シートと同槗。"
                        + " 有効ファイルがあれば read_excel 経路を省略可。"
                        + " JavaFX 環境変数タブの既定は"
                        + " code/exclude_rules.json（実在時）、無ければ code/json/stage1_exclude_rules.json（実在時）。"
                        + " JavaFX は「ファイル...」で選択可。");
        put(
                "PM_AI_PLAN_RESULT_TASK_JSON",
                "段階2 出力 production_plan_*.xlsx と同名ベースの"
                        + " 結果_タスク一覧.json（サイドカー）読み書き："
                        + "0/false/no/off/none で無効。有効時は再読込を"
                        + " JSON 優先にして Excel I/O を削減。");
        put(
                "PM_AI_PLAN_RESULT_TASK_JSON_PATH",
                "read_result_task_dataframe が読む JSON の絶対パス"
                        + "（実在ファイルのとき"
                        + " 出力 xlsx 横のサイドカーパスより優先）。");
        put(
                "PM_AI_STAGE2_WRITE_EXCEL",
                "段階2 で production_plan / member_schedule の xlsx を出力先に残すか。"
                        + " 0/false/no/off/none で JSON のみ（内部で一時 xlsx を生成し JSON"
                        + " 出力後に破棄）。"
                        + " 未設定または 1 で従来通り xlsx も出力。"
                        + " JavaFX の「実行・ログ」タブのチェックが段階2"
                        + " 起動時に本変数を上書きする。"
                        + " 0 のときは設備ガント（計画・実績明細）系シートは作成しない（処理時間の削減）。");
        put(
                "PM_AI_STAGE2_ENGINE",
                "段階2の実行エンジン。未設定・空・python（大小無視）で従来どおり Python 子プロセス（plan_simulation_stage2.py）。"
                        + " java のとき JVM 内の jp.co.pm.ai.planning.stage2 を起動し Python 段階2は使わない。"
                        + " 配台コアの完全な Python 同等は段階的に拡張する（現状は入力読取・最小成果物・JSON ミラーの足場）。"
                        + " 本番既定を java のみに切り替えるのは、JavaFX「Java/Python 同一検証」がチーム承認の golden 全件でパスした後に行う（Python 正本は比較・障害時用に残す）。");
        put(
                "PM_AI_STAGE2_GOLDEN_CI",
                "1 のときのみ JUnit `Stage2GoldenParityCiTest` が有効化される（通常の mvn test ではスキップ扱い）。"
                        + " `scripts/stage2_golden_parity_ci.sh` から設定して CI で段階2 Java 足場の回帰を追加する用途。");
        put(
                "PM_AI_STAGE2_HEADLESS_CI",
                "1 のときのみ JUnit `Stage2HeadlessParityCiTest` が有効化される（Python→Java の同一検証ランナーを CI で叩く）。"
                        + " `planning_core` を import できる Python（3.14+ 推奨）が必要。未設定または import 失敗時はテストをスキップする。"
                        + " `scripts/stage2_headless_parity_ci.sh` を参照。");
        put(
                "PM_AI_STAGE2_JAVA_DELEGATE_PYTHON_DISPATCH",
                "段階2で PM_AI_STAGE2_ENGINE=java のときのみ有効。1/true/on/yes のとき、JVM 内 PassThrough ではなく"
                        + " Python plan_simulation_stage2.py（_generate_plan_impl 正本）を子プロセスで実行し、"
                        + " Python エンジンと同一の計画／人員成果物を出す（完全 Java 移植までの本番同一出力用）。"
                        + " PM_AI_CODE_PYTHON_DIR・PM_AI_PYTHON 等は Python 子と同様に必須。");
        put(
                "PM_AI_XLWINGS_STAGE2_DISABLED",
                "1/true/yes/on で段階2後の xlwings"
                        + "（列設定シート図形複製等、Excel COM/アドイン連携用）"
                        + " をスキップ。openpyxl の xlsx 保存は從来通り。"
                        + "JavaFX からの段階2のみなら本条は実質無関係となることが多い。");
        put(
                "MASTER_WORKBOOK_FILE",
                "master.xlsm のファイル名（空で master.xlsm）。"
                        + "マクロブック階層からの相対パス可。"
                        + " PM_AI_MASTER_WORKBOOK 未指定時の解決に使用。"
                        + " 「マスタ読込サマリ」タブと連動。");
        put(
                "MASTER_USE_SPEED_SHEET",
                "master 内 speed シートによる加工速度上書きを有効化。");
        put(
                "STAGE2_DISPATCH_FLOW_TRIAL_ORDER_FIRST",
                "日内配台フロー: 1=試行順優先マルチパス（既定）、"
                        + "0=従来ソート。");
        put(
                "STAGE2_GLOBAL_DISPATCH_TRIAL_ORDER_STRICT",
                "配台試行順の「枠」より大きい順への割り込み制限。");
        put(
                "STAGE2_COPY_COLUMN_CONFIG_SHAPES_FROM_INPUT",
                "段階2後、列設定シートの図形を xlwings で複製"
                        + "（Excel アドイン/マクロ連携時。"
                        + "JavaFX での headless 段階2は通常関係なし）。");
        put(
                "PM_AI_CMD_PAUSE_ON_ERROR",
                "CLI 終了時の pause（Windows）。"
                        + "0/false で無効化（workbook_env_bootstrap 同様）。"
                        + "JavaFX デスクトップが起動する Python 子プロセスでは、環境タブの値に関わらず 0 に固定（pause によるハング防止）。");
        put(
                "PYTHONUTF8",
                "子プロセスで最終固定 1（本 UI では上書き不可）。");
        put(
                "PYTHONIOENCODING",
                "子プロセスで最終 utf-8 固定（本 UI では上書き不可）。");
        put(
                "XLWINGS_SUSPEND_AUTO_CALCULATION",
                "xlwings が Excel 書き込み前後で自動計算を手動に切替えるか"
                        + "（Excel アドイン連携時のみ意味がある。"
                        + "JavaFX から子プロセスで Excel を操作しない限り実質未使用）。");
        put(
                "PLAN_INPUT_DISPATCH_TRIAL_ORDER_LOCAL_ONLY",
                "配台試行順更新時に post_load（事後変形）をスキップ。");
        put(
                "TASK_PLAN_SHEET",
                "配台計画シート名（空で既定名）。");
        put(
                "STAGE2_SERIAL_DISPATCH_BY_TASK_ID",
                "日内配台: 1=依頼NO出現順で直列（他依頼は進まない）。");
        put(
                "PLANNING_B1_INSPECTION_EXCLUSIVE_MACHINE",
                "B-2/B-3: 熱融着検査の設備占有制御。");
        put(
                "PLANNING_B2_EC_FOLLOWER_DISJOINT_TEAMS",
                "B-2/B-3: EC と後続工程の担当者集合を分離。");
        put(
                "WIP_LIMIT_EC_BEFORE_INSP_ROLLS",
                "工程間 WIP: EC前〜検査までのロール上限。");
        put(
                "RAW_FABRIC_WIDTH_TABLE_PATH",
                "原反幅 CSV（planning_core の外部表参照）。");
        put(
                "PRODUCT_WIDTH_TABLE_PATH",
                "製品幅 CSV。空だとマクロブック階層で探索。");
        put(
                "COMPARE_GANTT_SNAPSHOT_DIR",
                "plan_compare_gantt_from_snapshot.py: 比較元の日時フォルダ"
                        + "（pdf 配下の最新を選択可）。");
    }

    private EnvVarDocs() {}

    private static void put(String key, String text) {
        LOGIC.put(key, text);
    }

    /** Logic-derived note, or empty if unknown. */
    public static String logicOnly(String key) {
        if (key == null || key.isBlank()) {
            return "";
        }
        return LOGIC.getOrDefault(key.trim(), "");
    }

    /**
     * Merges sheet/import description with {@link #logicOnly}; avoids duplicate when one contains the other.
     */
    public static String mergeDescriptions(String sheetDescription, String key) {
        String logic = logicOnly(key);
        String s = sheetDescription != null ? sheetDescription.trim() : "";
        if (s.isEmpty()) {
            return logic;
        }
        if (logic.isEmpty()) {
            return s;
        }
        if (s.contains(logic) || logic.contains(s)) {
            return s.length() >= logic.length() ? s : logic;
        }
        return s + " — " + logic;
    }
}
