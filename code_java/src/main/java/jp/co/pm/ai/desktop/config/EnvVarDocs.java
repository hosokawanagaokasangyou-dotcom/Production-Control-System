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
                        + "（1）pm-ai-package-release フォルダの UNC（推奨。湖南／国分は工場既定で切替）。直下の外付け version.txt と PMD_version_upgrade.zip を参照する。"
                        + "（2）PMD_version_upgrade.zip のフルパス（ZIP 隣の version.txt で版比較）。"
                        + "空のときは自動更新しない（情報表示のみ）。");
        put(
                "PM_AI_RDP_PORTABLE_BUNDLE_SOURCE_DIR",
                "リモートデスクトップ専用ポータブル（PmAiRpaLuncher.exe）の版アップ正本。"
                        + " ローカルビルド出力はリポジトリ直下 rpa_luncher_release。"
                        + " 湖南工場既定: M:\\湖南工場\\…\\共有DATA\\PmAiRpaLuncher_portable"
                        + "（PmAiRpaLuncher\\PmAiRpaLuncher.exe の2段上。version.txt + PmAiRpaLuncher_version_upgrade.zip）。"
                        + " 配台 PMD の正本は PM_AI_PORTABLE_BUNDLE_SOURCE_DIR（pm-ai-package-release）。");
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
                        + "配台不要は master.xlsm から PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK 同フォルダの stage1_exclude_rules.json に書き出し。"
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
                "PM_AI_ALADDIN_MASTER_DIR",
                "依頼書入力タブが参照するアラジンマスタフォルダ（フルパス）。"
                        + "後加工商品／加工内容／工程マスタ xlsx と"
                        + "マスタリレーション統合結果.xlsx を置く。"
                        + "空のときはグローバル設定の工場（湖南／国分）に応じた UNC 既定。"
                        + "湖南: 共有DATA/アラジンマスタ。国分: DATA/アラジンマスタ。"
                        + "環境変数を初期化（工場選択）でも上書きされる。"
                        + "環境変数タブのフォルダ選択可。"
                        + "統合スクリプト create_integrated_master.py にも渡される。");
        put(
                "PM_AI_POSTPROC_PRODUCT_MASTER_UPLOAD",
                "後加工商品マスタのアップロード用 xlsx（フルパス）。"
                        + "依頼書入力【設定】タブのマスタ編集カードが読み書きする。"
                        + "空のときは PM_AI_ALADDIN_MASTER_DIR 直下の"
                        + "アップロード用_後加工商品マスタ.xlsx。"
                        + "見出し行は参照の後加工商品マスタ.xlsx と同一である必要がある。");
        put(
                "PM_AI_REQUEST_FORM_JUCHU_FILE",
                "依頼書入力タブの受注データベース Excel（受注ﾌｧｲﾙ シート）のフルパス。"
                        + "例: 加工依頼書入力.xlsm。"
                        + "空のときは工場既定 UNC（湖南: 生産管理システム/管理システム/加工依頼書入力.xlsm、"
                        + "国分: 加工管理/加工計画関連/加工依頼書入力　国分コピー入力（2024年12月28））.xlsm）。"
                        + "環境変数初期化の工場選択でも設定。"
                        + "環境変数タブのファイル選択可。設定タブのパス表示・自動転記先に使用。");
        put(
                "PM_AI_REQUEST_FORM_ORIGINAL_DIR",
                "依頼書入力タブがスキャンする依頼書原本フォルダ（フルパス）。"
                        + "*加工依頼書*.xlsm 等の原本 Excel が置かれたディレクトリ。"
                        + "初期値は空（任意）。未設定のままでも起動可能。"
                        + "起動時に未設定なら BOX フォルダ選択を1回案内（スキップ可）。"
                        + "未設定時の実行時解決は工場既定（受注ファイル既定の親フォルダ）にフォールバック。"
                        + "環境変数タブのフォルダ選択可。"
                        + "プレビュー・照合・背景キャッシュの原本読込先。");
        put(
                "PM_AI_REQUEST_FORM_TPI_PDF_DIR",
                "TPI（東レペフ加工品）依頼書 PDF のスキャン先フォルダ（フルパス）。"
                        + "ECOWD/JR 系・後加工/PN 系の *.pdf を半自動取込する。"
                        + "空のとき湖南工場既定 UNC（共有DATA/TPI依頼書）。国分は空（手動指定）。"
                        + "依頼書入力タブの照合・PDF プレビュー・parse キャッシュに使用。"
                        + "フォルダ配下は読取専用（書込・削除禁止）。"
                        + "環境変数タブのフォルダ選択可。");
        put(
                "PM_AI_TESSERACT_CMD",
                "Tesseract OCR 実行ファイル（tesseract.exe）のフルパス。"
                        + "TPI 依頼書 PDF が画像スキャン（テキスト抽出不可）と判定されたとき OCR 読取に使用。"
                        + "空のとき C:\\Program Files\\Tesseract-OCR\\tesseract.exe 等を探索。"
                        + "環境変数タブのファイル選択可。");
        put(
                "PM_AI_TESSERACT_TESSDATA_DIR",
                "Tesseract 言語データ（tessdata）フォルダ。jpn.traineddata が必要。"
                        + "空のとき PM_AI_TESSERACT_CMD 近傍または Program Files\\Tesseract-OCR\\tessdata を探索。"
                        + "環境変数タブのフォルダ選択可。");
        put(
                "PM_AI_REQUEST_FORM_PREVIEW_PDF_CJK_SCALE",
                "依頼書 PDF プレビュー: 日本語フォントサイズ補正係数（Excel pt に乗算）。"
                        + "範囲 0.50～1.00。空または未設定で 0.72。"
                        + "PDF が Excel より大きくはみ出すときは 0.65～0.70、小さすぎるときは 0.80 前後。"
                        + "変更後はプレビュー PDF キャッシュが再生成される。");
        put(
                "PM_AI_REQUEST_FORM_RDP_PROFILE",
                "依頼書入力タブ「リモートデスクトップ」で起動する RDP プロファイル（.rdp）のフルパス。"
                        + "空のときは未設定。環境変数タブのファイル選択可。"
                        + "子タブの「リモートデスクトップを起動」で mstsc.exe に渡されます（Windows のみ）。");
        put(
                "PM_AI_RDP_COMPANION_PROGRAM",
                "RDP 接続先サーバー上で起動するプログラムのパス。"
                        + "「リモートデスクトップを起動」時に .rdp へ alternate shell（接続時の初期プログラム）として書き込む。"
                        + "例: C:\\Windows\\System32\\notepad.exe（接続先のパス）。空なら通常デスクトップ接続。");
        put(
                "PM_AI_RDP_COMPANION_PROGRAM_ARGS",
                "PM_AI_RDP_COMPANION_PROGRAM の引数。"
                        + " .rdp の alternate shell の引数として書き込む。空なら引数なし。");
        put(
                "PM_AI_RDP_FULLSCREEN",
                "RDP 起動時の全画面設定。1/true/on=全画面、0/false/off=ウィンドウ（既定）。"
                        + " ウィンドウ時は PM_AI_RDP_DESKTOP_WIDTH/HEIGHT を .rdp へ書き込み mstsc をウィンドウ表示する。"
                        + " 配台計画システムの背面で接続先 RPA を動かす用途では 0 を推奨。");
        put(
                "PM_AI_RDP_DESKTOP_WIDTH",
                "RDP ウィンドウ表示時の幅（ピクセル）。PM_AI_RDP_FULLSCREEN=0 のとき起動前に .rdp へ反映。既定 1920。");
        put(
                "PM_AI_RDP_DESKTOP_HEIGHT",
                "RDP ウィンドウ表示時の高さ（ピクセル）。PM_AI_RDP_FULLSCREEN=0 のとき起動前に .rdp へ反映。既定 1080。");
        put(
                "PM_AI_RDP_LAUNCHER_EXE",
                "接続先 RDP ランチャー（PmAiRdpRemoteLauncher.exe）のフルパス。"
                        + "空のときは配備先フォルダ（配台 PMD: PM_AI_RDP_LAUNCHER_DEPLOY_DIR、"
                        + "専用ランチャー: PM_AI_RPA_LAUNCHER_DEPLOY_DIR、各未設定時の既定は AppPaths 参照）の"
                        + " PmAiRdpRemoteLauncher.exe。");
        put(
                "PM_AI_RDP_LAUNCHER_INI",
                "接続先 RDP ランチャー設定 ini のフルパス上書き。"
                        + "空のときは配備先フォルダ（PmAiRdpRemoteLauncher.exe と同階層）の {操作者名}_RPA設定.ini"
                        + "（PM_AI_OPERATOR_USER / セッション操作者）。操作者未設定時は RPA設定.ini。"
                        + " レガシー RAP設定.ini / DATA 配下は読取フォールバックのみ。"
                        + " PmAiRdpRemoteLauncher.exe に操作者名を引数で渡すと同じ ini を参照する。");
        put(
                "PM_AI_RDP_LAUNCHER_DEPLOY_DIR",
                "配台 PMD のリモートデスクトップタブ向け: 接続先 RDP ランチャー exe（PmAiRdpRemoteLauncher.exe）と"
                        + " RDP起動プロファイル.json、RPA設定.ini の配備先共有フォルダ（UNC 可）。"
                        + " 空のときは "
                        + AppPaths.DEFAULT_PM_AI_RDP_LAUNCHER_DEPLOY_DIR
                        + "。"
                        + " リモートデスクトップ専用ランチャー（PmAiRpaLuncher.exe）では PM_AI_RPA_LAUNCHER_DEPLOY_DIR を使用。");
        put(
                "PM_AI_RPA_LAUNCHER_DEPLOY_DIR",
                "リモートデスクトップ専用ランチャー（PmAiRpaLuncher.exe）向け: 接続先 RDP ランチャー exe と"
                        + " RDP起動プロファイル.json、RPA設定.ini の配備先共有フォルダ（UNC 可）。"
                        + " 空のときは "
                        + AppPaths.DEFAULT_PM_AI_RPA_LAUNCHER_DEPLOY_DIR
                        + "。"
                        + " 配台 PMD とは別キー（PM_AI_RDP_LAUNCHER_DEPLOY_DIR）。");
        put(
                "PM_AI_RDP_OPERATOR_USERS_STORE_DIR",
                "配台 PMD のリモートデスクトップタブ向け: 操作者 bin 保存フォルダ（UNC 可）。"
                        + " 空のときは "
                        + AppPaths.DEFAULT_PM_AI_RDP_OPERATOR_USERS_STORE_DIR
                        + "（"
                        + AppPaths.RDP_LAUNCHER_OPERATOR_USERS_BIN
                        + " とバックアップ）。"
                        + " 専用ランチャーでは PM_AI_RPA_LAUNCHER_OPERATOR_USERS_STORE_DIR を使用。");
        put(
                "PM_AI_RPA_LAUNCHER_OPERATOR_USERS_STORE_DIR",
                "リモートデスクトップ専用ランチャー向けの操作者 bin 保存フォルダ（UNC 可）。"
                        + " 空のときは掲示板共有 "
                        + AppPaths.DEFAULT_PM_AI_RPA_LAUNCHER_OPERATOR_USERS_STORE_DIR
                        + "（"
                        + AppPaths.RDP_LAUNCHER_OPERATOR_USERS_BIN
                        + " とバックアップ）。"
                        + " 前回選択した操作者名は PC ローカルの last-rdp-launcher-operator.txt に保存。"
                        + " 配台 PMD とは別キー（PM_AI_RDP_OPERATOR_USERS_STORE_DIR）。");
        put(
                "PM_AI_OPERATOR_USER",
                "起動時に選択した操作者名。子プロセス env に載せる。"
                        + " RPA設定.ini の 操作者= 行と operator-aladdin-credentials.launcher.json と組み合わせ、"
                        + " C# ランチャーが Aladdin RPA 起動引数 --id / --password を付与する。"
                        + " 資格情報本体は factory-operator-users.bin（リモートデスクトップタブで編集）。");
        put(
                "PM_AI_FACTORY_SITE",
                "利用工場（KONAN / KOKUBU）。工場切替コンボ（実行タブ）に追従して自動更新。"
                        + " C# ランチャーが operator-aladdin-credentials.launcher.json"
                        + " 内の工場ブロックを選ぶときに参照するほか、"
                        + " 段階1の湖南工場限定ロジック（DISPATCHABLE_FROM_TIME_KONAN_STOCK 等）の判定にも使用（未設定時は KONAN）。");
        put(
                "DISPATCHABLE_FROM_TIME_KONAN_STOCK",
                "湖南工場（PM_AI_FACTORY_SITE=KONAN）かつ受注ファイル「在庫場所」列に「湖南」を含むタスクのみ:"
                        + " 原反投入日と同日の配台開始下限（HH:MM）。既定 09:30。"
                        + " 他タスクは従来通り DISPATCHABLE_FROM_TIME(12:45) を使用。");
        put(
                "PM_AI_RDP_LAUNCHER_AUTO_DEPLOY",
                "依頼書 UI から接続先共有フォルダへランチャー exe を自動再配備する。"
                        + "0/false/off で無効。既定は有効。");
        put(
                "PM_AI_RDP_LAUNCH_PROFILE_NUMBER",
                "リモートデスクトップタブで最後に使用した起動プロファイル番号（1～9）。"
                        + " 次回起動時の ComboBox 既定値。"
                        + " RPA設定.ini の起動プログラム番号（スロット）と同一。"
                        + " 名称・説明は共有フォルダの RDP起動プロファイル.json に保存。");
        put(
                "PM_AI_RDP_EMBED_STARTUP_IN_PROFILE",
                "1/true/on のときのみ PM_AI_RDP_COMPANION_PROGRAM を .rdp へ alternate shell として書込。"
                        + "既定は off（接続先はタスクスケジューラ + RPA設定.ini）。");
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
                "master 系 .xlsm / .xlsx の絶対パス（UNC 可）。"
                        + " planning_core 子プロセスの必須 env。"
                        + " ファイル名のみのときは planning cwd（通常 code/）相対。"
                        + " 空のとき resolveMasterWorkbookCandidate 等で補完。"
                        + " JavaFX の「マスタ読込サマリ」タブで内容を確認可。");
        put(
                "PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK",
                "利用工場と本パスが別工場を指すとき"
                        + " summaryAiDispatchXlsxPathForFactory が工場既定 UNC へ切替"
                        + "（操作者 bin・PDF・バックアップと整合）。"
                        + " 段階1（run_stage1_extract）は同フォルダへ"
                        + " stage1_exclude_rules.json を書出し"
                        + " PM_AI_EXCLUDE_RULES_JSON を自動設定。"
                        + " 特別ルール JSON は dispatch_special_rules/ 配下。"
                        + " 操作者 bin・PDF・バックアップ・remote_log 等の出力先フォルダの基準パス。");
        put(
                "PM_AI_REMOTE_LOG",
                "リモートサポート用ログ（サマリ Excel 同階層の remote_log/操作者/）への"
                        + "段階1／2／2.1 終了時スナップショット："
                        + "0/false/no/off/none で無効（空で有効）。"
                        + " 実行・ログタブ本文と code/log/execution_log.txt を3日世代管理。");
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
                        + "を PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK と同一フォルダの"
                        + " stage1_exclude_rules.json へ書き出し、本変数を"
                        + " その絶対パスに自動設定（子プロセス内）。"
                        + " 手動でも UTF-8 JSON（list または {\"rules\":[...]}、"
                        + " 列構造は設定シートと同槗。"
                        + " 有効ファイルがあれば read_excel 経路を省略可。"
                        + " JavaFX 環境変数タブの既定はサマリ Excel 同フォルダの"
                        + " stage1_exclude_rules.json（無ければ code/exclude_rules.json または"
                        + " code/json/stage1_exclude_rules.json から初回コピー）。"
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
                        + " JavaFX から段階2 起動時は常に 1 を渡す。"
                        + " 0 のときは設備ガント（計画・実績明細）系シートは作成しない（処理時間の削減）。");
        put(
                "PM_AI_STAGE2_SKIP_TODAY_DISPATCH",
                "1/true/yes/on のとき、データ抽出日（当日）の暦日には配台せず、計画開始日を翌暦日以降にずらす（段階2）。"
                        + " JavaFX「配台計画_タスク入力」タブのチェックが子プロセス起動時に上書きする。");
        put(
                "PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH",
                "1/true/yes/on のとき、実加工数が正の行（加工途中相当）を配台キューに入れない（当日完了と想定、段階2）。"
                        + " JavaFX 段階2 起動時は常に無効（0）。加工途中は「配台計画_タスク入力」タブの翌日配台量ダイアログで指定。");
        put(
                "PM_AI_SKIP_GEMINI_API",
                "1/true/yes/on のとき Gemini generateContent を呼ばない（開発用）。"
                        + " JavaFX 実行・ログタブ「その他」内のチェックが子プロセス起動時に上書きする。"
                        + " 本番運用では 0 または未設定を推奨。");
        put(
                "PM_AI_STAGE2_IN_PROGRESS_NEXT_DAY_DISPATCH_JSON",
                "段階2直前に JavaFX が書く UTF-8 JSON（加工途中行の翌日配台量 m）。"
                        + " build_task_queue_from_planning_df が実加工数>0 の行の配台量を上書きする。"
                        + " 未設定・ファイル無しのときはシートの配台使用残数量を用いる。");
        put(
                "PM_AI_STAGE2_ALADDIN_TODAY_EXCLUDE_NEXT_DAY_JSON",
                "段階2直前に JavaFX が書く UTF-8 JSON（アラジン当日配台行の翌日除外量 m）。"
                        + " 実加工数≦0 の行について、翌稼働日の配台割当上限から差し引く。"
                        + " 未設定・ファイル無しのときは除外なし。");
        put(
                "PM_AI_OVERTIME_SIMULATION_JSON",
                "段階2.1 残業/休出シミュ: ウィザード確定時に JavaFX が書く UTF-8 JSON。"
                        + " working_overrides（休日出勤○/グレー）と overtime_minutes（分）を段階2.1 フル再配台に適用する。"
                        + " master.xlsm は変更しない。成果物は output/stage21/ へ出力。");
        put(
                "PM_AI_STAGE2_1_OVERTIME",
                "1/true/yes/on のとき段階2.1（残業/休出シミュ）のフル再配台。output/stage21/ へ成果物を分離出力する。");
        put(
                "PM_AI_LEARNING_ARCHIVE_ENABLED",
                "1/true/yes/on（既定）のとき学習アーカイブの背景実行を有効化（将来の学習パイプライン用）。");
        put(
                "PM_AI_DISPATCH_LEARNING_ARCHIVE_SUBDIR",
                "サマリ Excel 同フォルダ内の学習アーカイブサブフォルダ名（既定 dispatch-learning-archive）。");
        put(
                "PM_AI_LEARNED_SPEED_ENABLED",
                "1/true/yes/on のとき実績由来学習速度を段階1/配台計画読込時に適用。"
                        + " 既定は off。JavaFX 実行・ログタブのチェックが子プロセス起動時に上書きする。");
        put(
                "PM_AI_LEARNED_SPEED_MIN_SAMPLES",
                "学習速度適用に必要な (工程名,機械名) 別最小観測数（既定 5）。");
        put(
                "PM_AI_LEARNED_SPEED_PERCENTILE",
                "学習速度のパーセンタイル（既定 50 = p50）。");
        put(
                "PM_AI_LEARNED_SPEED_HISTOGRAM_BIN_WIDTH",
                "速度ヒストグラムのビン幅 m/分（既定 1.0）。");
        put(
                "PM_AI_DEBUG_STAGE3_PLAN_ACTUAL_SINGLE_LINE",
                "配台計画手動修正タブ: 段階3試行後の日付セル表示。"
                        + " 1/true/yes/on で（段階3前）（段階3後）を1行（スペース区切り）。"
                        + " 0/false/no/off/none（既定）で2行（改行）。2行は固定行高44px・wrap-text なし。"
                        + " Spreadsheet の layout IOOBE 回避のため。");
        put(
                AppPaths.KEY_PM_AI_STAGE3_UI_VISIBLE,
                "段階3.0～3.2に関するUIを表示する。"
                        + " 1/true/yes/on で表示。0/false/no/off/none、空欄、未設定では非表示（既定）。"
                        + " 段階3のロジック・既存成果物・実行履歴は削除しない。");
        put(
                "PM_AI_STAGE2_ENGINE",
                "段階2の実行エンジン（互換用キー）。JavaFX 実行タブからの段階2は常に Python 子プロセス（plan_simulation_stage2.py）のみ。"
                        + " 未設定・空・python（大小無視）で従来どおり。java が指定されていても無視され Python が起動する（旧 JVM 段階2は撤去済み）。");
        put(
                "PM_AI_XLWINGS_STAGE2_DISABLED",
                "1/true/yes/on で段階2後の xlwings"
                        + "（列設定シート図形複製等、Excel COM/アドイン連携用）"
                        + " をスキップ。openpyxl の xlsx 保存は從来通り。"
                        + "JavaFX からの段階2のみなら本条は実質無関係となることが多い。");
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
                "配台計画シート名（空で タスク一覧。マクロブックは 配台計画_タスク入力）。");
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
                "原反幅 CSV。既定は PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK 同フォルダ（無ければ code/ から初回コピー）。");
        put(
                "PRODUCT_WIDTH_TABLE_PATH",
                "製品幅 CSV。既定はサマリ Excel 同フォルダ（無ければ code/ から初回コピー）。");
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
