---
name: ExcelからJavaFXへUI移行
overview: **主目的は UI を JavaFX に移すこと**であり、Python をすべて Java に書き換えることは求めない。**業務ロジック（配台・ETL・検証）は Excel／VBA／COM に依存せず、JavaFX＋Python のみで完結させる。** **最終成果物は Excel が開ける形式（主に `.xlsx`）で提供し、現場での確認・共有に使えること。** **フォールバックはアプリ／Python パイプライン内で完結**させ、**フォールバック不可能な場合は致命的エラーとして処理を中断**し、**原因と推奨対処法をユーザーへ明示**する（旧 Excel への値貼り・VBA 迂回は設計しない）。**UI は Java と親和性の高いオープンソース・コンポーネント（例: ControlsFX 等）を積極利用し、Excel のシート・セル操作に近いフォーマット／UX を優先する。** **取り込む OSS は、当プロジェクトの利用・配布において永続的に無償で続けられる条件のものに限定する（下記「OSS の永続無料」）。** **グラフ・可視化も視野に入れ、チャートは既定で JFreeChart を優先し、状況に応じて JavaFX Chart（`javafx.scene.chart`）を選択的に併用する。** **配台コア等の重い処理**については、運用後に計測し、ボトルネックのみ **Java への部分実装（JNI／サブプロセスからの置換／同一 JVM 内ライブラリ）**を検討する余地を残す。当面は Python（planning_core）を呼び出す構成が既定。新規に生成・追加する Java／JavaFX コードは **`code_java/`** に置き、既存の Excel・VBA・Python ラインが **`code/`** に残るよう **ツリーを分離**する。Power Query のデータの流れはプラン **「Power Query 正本」**に集約する。**実装着手前に固める各種取り決め（基盤・IPC・データ契約・ETL・画面・配布・QA・ドキュメント・運用）は §「実装前に固める取り決め（決定値）」を正とする。**
todos:
  - id: inventory-vba-python
    content: VBA エントリと cmd/xlwings 可否をマトリクス化（xlwings_console_runner・各 *.py から一覧）
    status: pending
  - id: bootstrap-javafx
    content: リポジトリ直下の code_java に Maven・JavaFX・ControlsFX・POI（必要時）・JFreeChart・Windows 配布方針でプロジェクトを追加（依存は永続無料方針で選定・固定、既存 code/ と分離）
    status: pending
  - id: oss-perpetual-free-audit
    content: 全依存の LICENSE／利用条件を確認し永続無料であることを記録（記録先は plan/oss_licenses.md の Markdown 表形式：コンポーネント名／バージョン／ライセンス種別／利用可否）。オープンコアの有償限定・サブスク必須・二重ライセンスで有償が必要な解釈は採用しない／法務確認
    status: pending
  - id: charts-jfreechart-javafx
    content: グラフ要件を棚卸しし JFreeChart を既定・JavaFX Chart を選択利用。SwingNode 埋め込み・画像エクスポート等は POC で決定（ライセンス・見た目）
    status: pending
  - id: mvp-python-bridge
    content: ProcessBuilder で TASK_INPUT_WORKBOOK 付与・段階1/2 実行・ログ表示の MVP
    status: pending
  - id: sheet-ui-parity
    content: 主要シート列定義の単一ソース化と OSS グリッド（例: ControlsFX SpreadsheetView）で編集・保存（openpyxl 互換）
    status: pending
  - id: powerquery-etl-parity
    content: プラン「Power Query 正本」PQ-A〜D を一次資料に ETL 代替を設計。共有フォルダ系は Python 再現＋ゴールデン比較。_q結果_配台表は直読＋型マニフェスト。ブックから M 全文エクスポートし正本と差分があれば正本を更新
    status: pending
  - id: pq-m-export-from-workbook
    content: UI参照ブックで各クエリの M を全文エクスポートし、プラン正本の「未確定」とユーザー貼付 M の差分を解消（クエリ正式名・ロード先シート・参照クエリの有無）
    status: pending
  - id: ui-reference-inventory
    content: plan/UI参照用_生産管理_AI配台(RC1).xlsx を基に画面・シート一覧・操作フローを棚卸しし JavaFX 画面一覧に落とす
    status: pending
  - id: xlwings-retirement
    content: xlwings 必須処理を Python ファイル I/O 経路へ寄せる／例外のみ別対応
    status: pending
  - id: decommission-excel-ui
    content: 運用切替後に VBA 役割を縮小・ドキュメント更新
    status: pending
  - id: hotspot-java-eval
    content: JavaFX＋Python 安定稼働後、プロファイルで重い処理を特定し、ROI が見える単位だけ Java モジュール化（ゴールデンデータ一致を前提）を検討
    status: pending
  - id: decision-l0-toolchain-finalize
    content: L0 ツールチェーンの最終固定（Windows 11 限定／JDK・JavaFX 最新 LTS のバージョン番号／Maven／code_java/ の pom.xml とサブモジュール構成）
    status: pending
  - id: decision-1-python-bridge-spec
    content: 1. Python↔JavaFX IPC 仕様書化と両側スタブ（§IPC テンプレートの type／フィールド名に準拠・ipc-line.schema.json 化・stdout/stderr 役割・終了コード 0/1/2/3/9・致命的時は原因・推奨対処を必須・.cancel フラグ・UI 単一実行）
    status: pending
  - id: decision-2-data-contract
    content: 2. データコントラクト整備（schema.json／settings.json／POI による .xlsm 完結／output 命名規則／LocalDateTime 境界変換。§IPC テンプレートの manifest 例と JSON Schema ファイルを一致させる）
    status: pending
  - id: ipc-json-schema-files
    content: code_java の resources/schema に ipc-line.schema.json（stdout NDJSON 1 行）・columns.schema.json（schema.json 検証）を配置し、プランのドラフトと同期
    status: completed
  - id: decision-3-etl-python-parity
    content: 3. PQ-A〜D の Python ETL 実装（fetch_latest_files 共通化、tests/fixtures ゴールデン、PQ-B は抽出時間最新で Drop Duplicates、ファイルの変換 (2)/(3) は M 抽出後に実装）
    status: pending
  - id: decision-4-screens-ia
    content: 4. 4 主要画面（設定／実績データ取込／計画データ編集（SpreadsheetView）／結果設備ガント）の単方向状態サイクル設計とワイヤー
    status: pending
  - id: decision-5-distribution
    content: 5. ZIP 配布＋run.bat（初回 pip install）／日付バージョニング v2026.MM.DD／資格情報は配布外（OS 環境変数または .env）／UNC 即時エラー／logs 14 日ローテ
    status: pending
  - id: decision-6-qa-test-bat
    content: 6. test.bat（Maven test＋pytest 一括）と許容誤差ルール（日付・文字列・整数は完全一致／小数は第2位丸め一致）。UI は手動 QA、TestFX 不採用
    status: pending
  - id: decision-7-docs-licenses
    content: 7. plan/oss_licenses.md（Markdown 表形式）と NOTICE.txt（配布同梱）。プラン同期は手作業（.cursor/plans → plan/ を上書きコミット）
    status: pending
  - id: decision-8-dual-running
    content: 8. Excel 論理切り離し・互換 xlsx 出力・フォールバックはロジック内のみ／不可能時は致命的エラーで中断し原因・推奨対処を UI 通知（旧 Excel 値貼り・VBA 迂回は不採用）。移行検証は新旧出力の比較に限定。配台パターンサマリは簡易フォーム→Python 引数
    status: pending
isProject: false
---

# Excel UI から JavaFX への大規模リファクタリング（プラン）

## 移行のスコープ（方針）

- **第一義**: **操作 UI を JavaFX に移す**（画面・入力・実行トリガ・ログ・**結果表示・グラフ／可視化**）。Excel を前面の操作端としない。
- **業務ロジックと Excel の切り離し**: **配台・ETL・検証の論理は Excel／VBA／COM／xlwings に依存しない**。入力・出力は **ファイル（`.xlsx`／設定 JSON 等）と明示パスのみ**とし、マクロ実行やブック常駐を前提にしない。
- **成果物の Excel 互換**: **最終出力は Excel が開いて閲覧・印刷・共有できる形式**（主に **`.xlsx`**）を提供する。これは **ファイル形式の互換**であり、**処理フローが Excel に依存する**ことを意味しない。
- **フォールバックと致命的エラー**: **フォールバックは JavaFX／Python のロジック内で完結**（縮退処理・代替経路・ユーザー指示による再試行）。**フォールバック不可能な場合は致命的エラーとして処理を中断**し、**原因と推奨対処法**（設定・ネットワーク・権限・入力データのどこを直すか）を **UI で明示**する。**旧 Excel への値貼り付けや VBA 実行への迂回は採用しない**。
- **既定の計算・入出力**: **`planning_core`（Python）をそのまま利用**する（子プロセス起動、環境変数、出力ファイル読込など）。**全面 Java 書き換えは対象外**とする（`_core.py` 規模・検証コストのため）。
- **オプション（第 2 段階以降）**: **プロファイルで時間・メモリが支配的な処理**（例: 日次ループ割付、巨大 DataFrame 変換、PQ 相当 ETL のうち頻繁に走る部分など）を特定し、**Java で高速化したモジュールに差し替える**ことを **視野に入れる**。その際は **入出力形式を変えず**、同一データで **Python 版と結果一致（許容誤差を定義）**を確認してから切り替える。
- **選ばないと決めるまで Java 化しない**: 「Java の方が速いから」だけでは移植しない。**計測根拠**と **保守コスト（ロジック二重管理）**のトレードオフを記録してから着手する。
- **生成コードの配置**: Java／JavaFX で新規に追加するソース・ビルドスクリプト・モジュール定義は **`code_java/`** を既定ルートとする。**`code/` は既存の Excel・VBA・Python（planning_core）ライン**のままとし、同一ツリーへの混在やパス衝突を避ける。
- **UI は OSS を優先し Excel に寄せる**: **オープンソース**かつ **Java／JavaFX と親和性が高い**ライブラリを **積極採用**し、**Excel のグリッド・セル編集に近い操作感**（行列・見出し・セル単位の入力など）を設計目標とする。**採用する OSS は必ず「永続無料」で利用できるものに限定**する（下記「OSS の永続無料」）。**グラフ化も視野**に入れ、チャート系は下記「UI 技術方針」のとおり **JFreeChart を優先**する。詳細は下記「UI 技術方針」。

## 現状の把握（リポジトリ上の事実）

- **Java / JavaFX のコードベースは未存在**（`pom.xml` / `build.gradle` / `*.java` なし）。プロジェクト名は「JAVA」だが、実装の正は **Python**（[code/python/planning_core](code/python/planning_core)）と **Excel ブック**に集約されている。
- **配台ロジック**は `planning_core`（特に巨大な [code/python/planning_core/_core.py](code/python/planning_core/_core.py)）にあり、**VBA は主に起動・環境変数・xlwings／cmd 経由の子プロセス制御**を担う（例: [code/python/xlwings_console_runner.py](code/python/xlwings_console_runner.py) の `run_stage1_for_xlwings` / `run_stage2_for_xlwings` 等）。
- **データの入出力**は `TASK_INPUT_WORKBOOK` 環境変数で指す **マクロブック（.xlsm）**と **マスタ（例: `master.xlsm`）**のシート列と強く結合している（[code/python/planning_core/__init__.py](code/python/planning_core/__init__.py) 冒頭のパッケージ doc、[`workbook_env_bootstrap`](code/python/workbook_env_bootstrap.py) の「設定_環境変数」シート）。
- **基礎データ（マスタ）の提供形態**: **Excel ブック（`.xlsx` / `.xlsm`）として提供する**。編集・更新は **Excel を主とした運用**とし、JavaFX／Python 側は **openpyxl／POI 等のファイル読込のみ**（Excel プロセス・マクロ実行は不要）。リポジトリ上の参照・ひな形の例: **[plan/master.xlsx](plan/master.xlsx)**（運用時の実ファイル名・パスは `MASTER_WORKBOOK_FILE` 等の設定と整合）。

### リポジトリ配置（`code` と `code_java` の役割分担）

| パス | 役割 |
|------|------|
| **`code/`** | **既存 Excel 版ライン**。VBA、`code/python`（planning_core）、要件定義ドキュメント等。従来どおりここを既存資産の正とする。 |
| **`code_java/`** | **JavaFX／Java で新規生成・追加するコード**の既定ルート（アプリ本体・Maven・将来の選択的 Java モジュール）。**`code/` との混在・衝突を避ける。** |

CI・IDE のワークスペース設定でも **両ツリーを明示的に分ける**（例: Java プロジェクトのルートは `code_java` のみ）。

要件定義 HTML でも「Excel を前面の操作画面にしない」移行先の例として **JavaFX** が言及されている（[code/要件定義/工程管理AI配台システム_経営層向け説明.html](code/要件定義/工程管理AI配台システム_経営層向け説明.html) 付近）。

### UI 参照用ブック（設計時のたたき台）

- リポジトリ内の **[plan/UI参照用_生産管理_AI配台(RC1).xlsx](plan/UI参照用_生産管理_AI配台(RC1).xlsx)** を、**現行 UI のレイアウト・シート構成・操作イメージの参照**として使う（JavaFX の画面ワイヤー／画面一覧の起点）。
- 実装フェーズでは、このブックから **画面単位の対応表**（シート／ボタン／一覧 → JavaFX 画面・コントロール）を切る。

### マスタ（基礎データ）ブック

- **提供形式**: **Excel（`.xlsx` / `.xlsm`）で提供する**。勤怠・スキル・マスタシート等の **基礎データの正本は Excel ファイル上のシート列**とする。
- **アプリ側の扱い**: **読み書きはファイル API のみ**（§「Excel との論理的切り離し」と両立。**データの体裁は Excel**、**実行時ロジックは Excel に依存しない**）。
- **参照**: プラン配下の **[plan/master.xlsx](plan/master.xlsx)** を、レイアウト・シート構成の参照例として用いる（運用ブック名は `master.xlsm` 等、環境設定に従う）。

### UI 技術方針（Java 親和の OSS・Excel に近いフォーマット）

| 観点 | 方針 |
|------|------|
| **OSS の永続無料（必須）** | **オープンソースを利用するときは、必ず永続的に無償で利用できる条件のものに限定する。** 具体的には次を満たすこと。**(1)** 当方の **構築・運用・配布**に対し、**継続利用のために別途ライセンス購入・サブスクリプションが必須**となる製品（またはその機能）のみに依存しない。**(2)** **オープンコア**で実質コア機能が有償版限定、**試用期間後の有償化**、**外部 SaaS の従量課金が前提**など、**無償利用が途切れる設計**は採用しない。**(3)** **二重ライセンス**（コミュニティ版と商用版）のうち、当プロジェクトの利用形態で **有償ライセンスが必要**になる解釈のものは、**採用前に法務確認**し、不可なら **代替 OSS** に切り替える。**(4)** **Apache 2.0 / BSD / MIT** 等の寛容ライセンスを優先。**LGPL / GPL** 系は **コンプライアンス手順（組み合わせ・改変・配布）が組織で許容される場合のみ**採用（例: JFreeChart）。**依存ごとに LICENSE を確認**し、TODO **`oss-perpetual-free-audit`** で一覧化する。 |
| **基本姿勢** | **オープンソース**で **Gradle／Maven から取り込みやすく**、**JavaFX 標準 API と共存しやすい**ものを優先。上記 **永続無料**を満たすことを前提に、同等機能なら **保守・ライセンス・実績**で選ぶ。 |
| **グリッド（Excel ライク）** | 既定の検討対象として **ControlsFX** の **SpreadsheetView**（セル単位編集・結合・フィルタ等の土台）を **POC の起点**とする。単純一覧のみは **TableView** でよいが、**業務シート相当はスプレッドシート系コンポーネントを優先**。 |
| **書式・ファイル整合** | **`Apache POI`** を **読み書き・書式検証**に活用し、Python **openpyxl** との列・互換を確認できるようにする（同一 JVM 内で検証しやすい）。 |
| **グラフ・チャート** | **グラフ化を設計スコープに含める**。**既定で JFreeChart を優先**（業務チャートの表現力・実績）。**JavaFX 標準の Chart**（`javafx.scene.chart`、折れ線・棒・積み上げ等）は **軽量ダッシュ・JavaFX レイアウトとの一体感が重要な画面**で **状況に応じて選択的に併用**する。同一アプリ内での **使い分け基準**（例: 複雑な複合チャート・既存レポート近似は JFreeChart、単純トレンドは JavaFX Chart）を POC で文書化。**JFreeChart の JavaFX 埋め込み**は **`SwingNode`** で Swing パネルとして載せる、または **画像（`BufferedImage`）生成して `ImageView`** に渡す等から POC で決定。ライセンス（JFreeChart は **LGPL** 系のため組織ポリシーで要確認）・メンテ版を固定する。 |
| **その他（必要時）** | 複雑インセル編集なら **RichTextFX**、アイコンなら **Ikonli** 等（OSS・Apache 2.0 系が多い）を **要件が出た時点で**追加評価。 |
| **検証** | UI 参照ブックと並べて **見た目・操作のギャップ**を POC で確認してから全面適用。チャートは **Excel 既存グラフとの近似度**も評価項目に含める。 |

## Power Query 正本（チャットでユーザーが貼付した M のみ）

### 文書化ルール（ハルシネーション防止）

- **リポジトリ格納の M 全文**: ユーザーがチャットで貼付した M は **[plan/power-query-m-sources.md](plan/power-query-m-sources.md)** の一覧どおり、`plan/01_*.m` … `plan/04_*.m` に保存している（トランスクリプトから抽出）。ブック内の最新 M と差分がある場合は **Excel 側を正**として `.m` を更新する。
- **出典の範囲**: 本節の **入力パス・関数名・列名・制御フロー**は、**本チャット内でユーザーが貼付した Power Query (M) 断片に現れる記述だけ**を根拠にする。リポジトリ内の `.xlsx` / `.xlsm` をこのプラン作成時点では開いて検証していない。
- **プランに書かないこと**: ブック内の **クエリの正式名称（Excel のクエリ一覧）とシート名の 1 対 1**、**「読み込み先テーブル」以外のロード設定**、ユーザーが M に含めていない **追加の変換ステップ**は **推測で補わない**。必要なら **[plan/UI参照用_生産管理_AI配台(RC1).xlsx](plan/UI参照用_生産管理_AI配台(RC1).xlsx)** を開き、Power Query エディタで **M をエクスポートして正本へ追記**する運用とする。
- **参照テーブルについて**: Power Query の **「参照」クエリ**や **最終的にワークシートに展開されるテーブル名**は、貼付 M だけでは特定できない場合がある。その場合は下表の「参照テーブル／ロード先」欄を **未確定（ブック要確認）** とする。

### 総覧（ユーザーが明示したクエリ名・入力・出力イメージ）

| ID | ユーザーが示した呼び名 | 入力の起点（M より） | 「最新」判定などファイル選択 | 出力側で M に現れる名前 |
|----|------------------------|----------------------|------------------------------|---------------------------|
| **PQ-A** | チャット上は「加工計画DATA」として説明。**PQ エディタ上のクエリ名はチャット未記載** | `Folder.Files("\\192.168.0.101\...\●DATA\生産計画問合せ")` | 非表示除外 → `Date created` **降順** → **先頭 1 件** `{0}` | 列 **`抽出時間`**（値は選ばれたファイルの `Date created`）。最終ステップ名は M 上 `抽出時間列の追加` |
| **PQ-B** | **`_q加工計画DATA_実績比較用`** | **PQ-A と同一 UNC** の `Folder.Files(...生産計画問合せ)` | `Date created` **降順** → **`Table.FirstN(..., 20)`**（コメントは「最新の8ファイル」とあるが **数値は 20**）→ 非表示除外 | 各ファイル処理後に列 **`抽出時間`**（各ファイルの `Date created`）。全体を **`Table.Combine`** 後 **`抽出時間` 降順ソート**。コメントに「重複の削除（Upsert）」があるが **提示 `in` はソートのみ** |
| **PQ-C** | **`_q加工実績明細DATA`** | `Folder.Files("\\192.168.0.101\...\002  加工G\●検査表作成\加工実績明細DATA")` | `Date accessed` **降順** → **`Table.FirstN(..., 1)`** → 非表示除外 | 変数 **`ファイル更新日時`**（先頭行の `Date modified`）。最終列 **`データ抽出時間`**（全行 `ファイル更新日時`）。列 **`累積実績`**・**`累積完了率`** 等 |
| **PQ-D** | **`_q結果_配台表`** | **`Excel.CurrentWorkbook(){[Name="フォルダパス"]}[Content]{0}[Column1]`** と **`結果_配台表.xlsx`** の連結 | 対象は **単一ファイル**（フォルダ走査なし） | `Excel.Workbook(...)` から **`Item="_t結果_配台表"` かつ `Kind="Table"`** の **`[Data]`**。続けて **`変更された型`** |

上表の **PQ-A〜D** の詳細ステップは以下に **M の出現順に近い形**で列挙する。

---

### PQ-A — 加工計画DATA（単一ファイル・生産計画問合せ）

**参照テーブル／シートへのロード先**: ユーザー提示 M には **クエリ名・シート名・「参照」作成の有無は含まれない** → **未確定（ブック要確認）**。

**ソース → 1 ファイルの決定**

1. `Folder.Files(パス)` … パス文字列はユーザー M の `パス = "\\\\192.168.0.101\\共有フォルダ\\湖南工場\\湖南共有\\生産管理システム\\管理システム\\●DATA\\生産計画問合せ"`（エスケープ表記は M 原文準拠）。
2. `Table.SelectRows(..., each [Attributes]?[Hidden]? <> true)`
3. `Table.Sort(..., {{"Date created", Order.Descending}})`
4. `最新ファイル = 並べ替えられたファイル{0}`
5. `現在の抽出時間 = 最新ファイル[Date created]`

**バイナリ → テーブル**

6. `現在のテーブル = ファイルの変換(最新ファイル[Content])` … **組み込みではなくブック内の同名関数**に依存。
7. `Table.Skip(現在のテーブル, 3)`
8. 全列について空文字 `""` を `null` に（`Table.TransformColumns` + `List.Transform(全列名, ...)`）

**見出し整形（ヘッダー領域のみ FillUp）**

9. `ヘッダーエリア = Table.FirstN(置き換えられた値, 2)`、`データエリア = Table.Skip(置き換えられた値, 2)`
10. `限定フィル = Table.FillUp(ヘッダーエリア, 全列名)` … FillUp の列リストは **`全列名`**
11. `結合 = Table.Combine({限定フィル, データエリア})`
12. `昇格されたヘッダー数 = Table.PromoteHeaders(結合, [PromoteAllScalars=true])`

**列名の再計算（`List.Buffer` + `List.Transform`）**

13. `旧ヘッダー`、`先頭行レコード = Table.First(昇格されたヘッダー数)`
14. `基準日` … `受注日` 列から有効シリアルの最小 → `#date(1899,12,30)` 加算、無ければ `Date.From(現在の抽出時間)`。`基準年`・`基準月` を派生。
15. `新ヘッダーリスト` … `_` 分割・末尾連番判定・先頭行との結合・`/` 日付の年補正（11–2 月と基準月の組み合わせで年 ±1）等（ユーザー M の `新見出し` ロジック全文に準拠）。
16. `見出しの更新 = Table.RenameColumns(..., List.Zip({旧ヘッダー, 新ヘッダーリスト}))`
17. `不要な先頭行を削除` … `先頭行レコード = null` でなければ `Table.Skip(..., 1)`

**仕上げ**

18. `回答納期` の `""` → `null`
19. `Table.TransformColumnTypes` … `受注日`,`原反投入日`,`出荷日`,`指定納期`,`回答納期`,`加工開始日`,`加工完了日` を `type date`
20. `Table.SelectRows(..., each [依頼NO] <> null and [依頼NO] <> "")`
21. 列名に `_加工速度` または `_加工時間` を **含む**列を `Table.RemoveColumns`
22. `抽出時間列の追加` … 列名 **`抽出時間`**、各行 `現在の抽出時間`、`type datetime`

---

### PQ-B — `_q加工計画DATA_実績比較用`（同一フォルダ・複数ファイル）

**参照テーブル／ロード先**: M 内にクエリ名 `_q加工計画DATA_実績比較用` は **チャットでユーザーが明示**。ワークシート名との対応は **未確定（ブック要確認）**。Python 側には **`加工計画DATA_実績比較用`** シート名がコード定数として存在する（[code/python/planning_core/_core.py](code/python/planning_core/_core.py) の `TASKS_SHEET_NAME_FOR_ACTUAL_GANTT_PLAN`）が、**PQ のロード先シート名と同一かはこのプランでは検証していない**。

**ソース → 最大 20 ファイル**

1. `Folder.Files` … **PQ-A と同一の生産計画問合せパス**
2. `Table.Sort(..., {{"Date created", Order.Descending}})`
3. `対象ファイル = Table.FirstN(並べ替えられたファイル, 20)` … **コメント「最新の8ファイル」と数値 20 の不一致はユーザー M 内の事実として記録**
4. 非表示除外

**各行（各ファイル）の内側クエリ**

5. `現在の抽出時間 = [Date created]`（行コンテキスト）
6. `現在のテーブル = ファイルの変換 (3)([Content])` … **PQ-A の `ファイルの変換` とは別名**
7. `Table.Skip(..., 3)`
8. `Table.ReplaceValue` … 列集合に **`倉庫       : 520201 湖南工場01本倉庫`** および **`Column2`〜`Column43`** を列挙（ユーザー M 原文）
9. `ヘッダーエリア` / `データエリア` に分割後、`Table.FillUp(ヘッダーエリア, フィル対象列)` … **`フィル対象列` は上記倉庫列＋Column2〜43 の明示リスト**（PQ-A の「全列名」と異なる）
10. `結合` → `PromoteHeaders`
11. `新ヘッダー = List.Transform(旧ヘッダー, each ...)` … PQ-A と同系の見出し変換（ユーザー M に **`List.Buffer` は無い**）
12. `見出しの更新`、`不要な先頭行を削除`、`回答納期` クリーンアップ、日付型、`依頼NO` フィルタ、`_加工速度`/`_加工時間` 列削除
13. **`抽出時間列の追加`** … 各ファイルの **`現在の抽出時間`**

**結合・終端**

14. `結合されたテーブル = Table.Combine(個別処理[処理済みテーブル])`
15. `並べ替えられた全データ = Table.Sort(結合されたテーブル, {{"抽出時間", Order.Descending}})` … **`in` はこれのみ**。コメントの「Upsert」と実装の差は **未解決フラグ**。

---

### PQ-C — `_q加工実績明細DATA`

**参照テーブル／ロード先**: クエリ名はユーザー明示。ロード先シートは **未確定（ブック要確認）**。

**ソース → 1 ファイル**

1. `Folder.Files("\\192.168.0.101\...\002  加工G\●検査表作成\加工実績明細DATA")`
2. `Table.Sort(..., {{"Date accessed", Order.Descending}})`
3. `保存された先頭行 = Table.FirstN(並べ替えられた行, 1)`
4. `ファイル更新日時 = 保存された先頭行{0}[Date modified]`
5. 非表示除外後、`ファイルの変換 (2)([Content])` 等で展開（ユーザー M の `RenameColumns`・`ExpandTableColumn` 連鎖に従う）
6. `Data` の展開列として **`Column1`〜`Column271`** を列挙（ユーザー M 原文どおり）
7. `Table.SelectRows(..., each [Column4] <> null)`
8. `Table.PromoteHeaders`

**中間変換・フィルタ**

9. 日付列の `TransformColumnTypes`
10. `Table.SelectColumns` で **検査備考を含めない列集合**（ユーザー M の長い列リスト）
11. 導出列 **`加工開始日時`**・**`加工終了日時`**・**`加工時間`**
12. `Table.SelectRows(..., each ([#"製造条件(内訳)"] = "長さ"))`
13. 列削減、`停機時間分(変換後)`、`加工開始日時(停機時間加算後)`、`Table.ReorderColumns`

**集計・結合**

14. `加工日付`、`日次集計`（`Table.Group` で `日次実績 = List.Max([実加工数])`）
15. `累積計算`（`依頼NO`・`工程名` 単位で日付昇順インデックス付き累積）
16. `Table.NestedJoin` / `ExpandTableColumn` で **`累積実績`** を明細へ
17. **`累積完了率`** … 換算数量が null または 0 なら 0、否则 累積実績÷換算数量、`Percentage.Type`
18. `加工日付` 列削除、**`データ抽出時間`** 追加（`each ファイル更新日時`）
19. 最終 `ReorderColumns`（ユーザー M の列順そのまま）

---

### PQ-D — `_q結果_配台表`

**参照テーブル／ロード先**: 取り込みオブジェクトは M 上 **`Item="_t結果_配台表"` かつ `Kind="Table"`**。ブック上のテーブル表示名とシート配置は **未確定（ブック要確認）**。

**ソース**

1. `現在のパス = Excel.CurrentWorkbook(){[Name="フォルダパス"]}[Content]{0}[Column1]`
2. `ファイルパス = 現在のパス & "結果_配台表.xlsx"`
3. `ソース = Excel.Workbook(File.Contents(ファイルパス), null, true)`
4. `_t結果_配台表_Table = ソース{[Item="_t結果_配台表", Kind="Table"]}[Data]`

**型の固定（ユーザー M に列挙されたそのまま）**

5. `変更された型 = Table.TransformColumnTypes(_t結果_配台表_Table, { ... })` … 列名・型の組は **ユーザー貼付 M のリスト全文と一致**する必要がある（プラン本文への再掲は冗長なため、**実装・検証時はチャット原文またはブックからエクスポートした M を一次資料とする**）。

---

### 本ドキュメントの保守手順

1. Excel で **[UI 参照ブック](plan/UI参照用_生産管理_AI配台(RC1).xlsx)** を開き、該当クエリの **詳細エディタで M を全文コピー**する。
2. 本プランの **PQ-A〜D** を、差分があれば **コピーした M にのみ基づき更新**する（推測で列名やステップを埋めない）。
3. 「参照テーブル」「ロード先シート」が分かった時点で、総覧表と各節の **未確定** を書き換える。

---

### 加工計画DATA と Power Query（データ供給の正）

「加工計画DATA」はブック内入力というより、**Power Query（M）で外部から生成**されている。JavaFX に UI を移すと **Excel 上の PQ 更新ボタンだけでは運用できない**ため、次をプランの前提に含める。

**詳細な変換ステップ**は **§ Power Query 正本の PQ-A** を一次記述とする（以下は移行方針の要約のみ）。

**データソース（要約・PQ-A と同一）**

- **フォルダ（UNC）**: `\\192.168.0.101\共有フォルダ\湖南工場\湖南共有\生産管理システム\管理システム\●DATA\生産計画問合せ`
- **ファイル選択**: 非表示以外を対象に **`Date created` 降順で先頭 1 ファイル**＝最新のみ使用。
- **変換の要点**: 先頭スキップ・空文字→null・ヘッダー／先頭行の結合と **`PromoteHeaders`**・列名の **`_連番` 除去と先頭行値との結合**・日付っぽい列名の **年跨ぎ補正（基準日は受注日列の最小シリアル or ファイル作成日時）**・型設定・`依頼NO` が空でない行のみ・列名に `_加工速度` / `_加工時間` を含む列削除・**`抽出時間`** 列に **現在選んだファイルの `Date created`** を付与。

**移行時の代替アーキテクチャ（どれかまたは併用を決める）**

| 方針 | 内容 | メリット／留意 |
|------|------|----------------|
| **PQ ロジックを Python に移植** | 同一 UNC から最新ファイルを選び、pandas で上記変換を再現し、`planning_core` が読む形式（シート相当の CSV または xlsx へ書き出し）にする | JavaFX から「データ取得実行」ボタンで完結可能。**M と Python の仕様一致検証**が必須 |
| **バッチ／タスクスケジューラ** | 定期的に変換済みファイルを共有フォルダの別パスへ出力し、アプリはそれを読む | Excel 不要。**鮮度と失敗時の通知**を運用設計 |
| **Excel 自動化で PQ のみ実行** | COM でブックを開きデータの更新のみ実行し保存（UI は JavaFX） | 実装は速い可能性。**Excel ライセンス・サーバ常駐・不安定さ**のリスク |
| **当面ハイブリッド** | 加工計画DATA の取得だけ Excel 側で更新→その後 JavaFX で配台のみ | 移行初期のリスク低減。**二重運用期間**が長くなりがち |

フェーズ 0 の棚卸しに **「PQ に依存するシート一覧」**と **「刷新後の単一取得パイプライン」**を追加する（**複数 UNC・複数 M パイプライン**を前提にする）。

### _q加工計画DATA_実績比較用 と Power Query（同一フォルダ・複数ファイル縦結合）

**ステップの正本**: § **PQ-B**。クエリ名 `_q加工計画DATA_実績比較用` はユーザーがチャットで明示。Python コードには **`加工計画DATA_実績比較用`** というシート名定数がある（[planning_core `_core.py`](code/python/planning_core/_core.py) の `TASKS_SHEET_NAME_FOR_ACTUAL_GANTT_PLAN`）が、**PQ のロード先シート名と同一かはブックで未検証**。

**移行上の示唆**

- PQ-A と **共通化できる処理**と、**N ファイル Combine・複数 `抽出時間`** の **分離**をコードで明示する。
- **Upsert** コメントと **実際の `in`（ソートのみ）** の差は ETL でも **未解決のまま踏まない**（§ PQ-B 参照）。

### _q加工実績明細DATA と Power Query（加工実績系）

**ステップの正本**: § **PQ-C**。ファイル選択は **`Date accessed`** ベースで **PQ-A/B と異なる**（総覧表）。

**移行上の示唆**

- **累積・日次 Max** など集計ロジックが重い。**pandas 移植時は PQ 出力とのゴールデン比較**が必須。

### _q結果_配台表 と Power Query（Python 出力 xlsx の読み戻し）

**ステップの正本**: § **PQ-D**。列型の一覧は **ユーザー貼付 M にそのまま存在**するため、実装時は **チャット原文またはブックからの M エクスポート**を一次資料とする。

**移行上の示唆**

- フォルダ走査が無く **出力ファイル直読で代替しやすい**。名前 **`フォルダパス`** は **アプリ設定へ移行**。
- **`planning_core` が `結果_配台表.xlsx` を出力する**ことはプラン上の **推定**に過ぎない（コードで確認したら正本に脚注を追加する）。

## 目標アーキテクチャ（推奨）

```mermaid
flowchart LR
  subgraph pq [現状_PQ供給_複数]
    Unc1[UNC_生産計画問合せ]
    Unc2[UNC_加工実績明細DATA]
    PQ1a[PQ_加工計画_単一ファイル]
    PQ1b[PQ_実績比較用_Nファイル結合]
    PQ2[PQ_実績明細]
    PQ3[PQ_結果_配台表]
    Named[名前_フォルダパス]
    OutFile[結果_配台表_xlsx]
    SheetA[加工計画DATA]
    SheetB[加工計画DATA_実績比較用]
    Sheet2[加工実績明細等]
    SheetH[_q結果_配台表表示]
    Unc1 --> PQ1a --> SheetA
    Unc1 --> PQ1b --> SheetB
    Unc2 --> PQ2 --> Sheet2
    Named --> PQ3
    OutFile --> PQ3
    PQ3 --> SheetH
  end
  subgraph before [現状_UIと計算]
    ExcelUI[Excel_VBA_UI]
    PyCore[planning_core_Python]
    Xlsx[マクロブック_xlsm]
    SheetA --> Xlsx
    SheetB --> Xlsx
    Sheet2 --> Xlsx
    SheetH --> Xlsx
    ExcelUI -->|RunPython_cmd_xlwings| PyCore
    PyCore --> OutFile
    PyCore --> Xlsx
  end
  subgraph after [移行後_推奨]
    Etl1[ETL_単一計画]
    Etl1b[ETL_実績比較結合]
    Etl2[ETL_実績明細]
    JavaFX_UI[JavaFX_UI]
    PyCore2[planning_core_継続]
    Xlsx2[ブックまたは中間形式]
    Out2[結果_配台表_xlsx]
    Unc1 -.-> Etl1
    Unc1 -.-> Etl1b
    Unc2 -.-> Etl2
    Etl1 --> Xlsx2
    Etl1b --> Xlsx2
    Etl2 --> Xlsx2
    JavaFX_UI -->|取得実行| Etl1
    JavaFX_UI -->|取得実行| Etl1b
    JavaFX_UI -->|取得実行| Etl2
    JavaFX_UI -->|Process_TASK_INPUT_WORKBOOK| PyCore2
    PyCore2 --> Out2
    PyCore2 --> Xlsx2
    JavaFX_UI -->|直接読込| Out2
  end
```

上図の **シート名・ノードラベル**は読みやすさのための略称。**PQ が実際にロードするテーブル／シート名は Excel ブック設定による**ため、確定情報は **§ Power Query 正本の総覧表と「未確定」欄**および、ブックからエクスポートした M を正とする。

- **UI 層**: Excel／VBA → **JavaFX**（デスクトップ）。これが本プロジェクトの主成果物。
- **ドメイン／計算（既定）**: **Python `planning_core` を子プロセスまたは将来は同一マシン上のサービスとして呼び出す**（`generate_plan` / `run_stage1_extract` 等）。**全面 Java 移植はスコープ外**。
- **ドメイン／計算（選択的・後追い）**: ボトルネックが **計測で特定できた場合のみ**、該当サブ問題を **Java ライブラリ化**し、JavaFX から直接呼ぶ／Python を薄いオーケストレーションに縮小する、などのハイブリッドを検討。**`_core.py` 全文移植は最終手段**とし、通常は **関数単位・パイプライン単位**が上限のイメージ。
- **永続化**: 移行初期は **既存の .xlsm/.xlsx シート構造を維持**し、Python がそのまま読める状態にするのがコスト最小。中長期で設定のみ JSON 化・DB 化するかは別判断。

## 主要な技術的ギャップ（必ず計画に含める）

1. **xlwings 依存**  
   一部処理は **Excel が開いた状態の COM（xlwings）**を前提にしている。JavaFX 単体では同等の「caller ブック」がないため、次のいずれかが必要になる。
   - **A**: 該当機能を **ファイルベースのみ**で完結するよう Python 側に経路を整理する（既に cmd 起動の `*.py` があるものは流用しやすい）。
   - **B**: 移行後も **Excel をバックグラウンドで起動**し xlwings を使う（運用・インストール複雑度は高い）。
   - **C**: Apache POI 等で Java から直接 xlsx を書き換え、Python はファイルのみ見る（二重実装リスクあり）。

   **移行後の目標**: **A のみ**とし、**ロジックが Excel ランタイムに依存しない**ことを検証する。**B は本プロジェクトの「Excel 論理切り離し」方針により採用しない**。実務的には **A を優先**し、どうしても必要な箇所だけ C を検討するのが安全。

2. **マクロブックの役割の分解**  
   現状は「UI（シート）」「設定（設定_環境変数 等）」「Python へのパス受け渡し」が同一 `.xlsm` に混在。JavaFX 化では **「プロジェクトディレクトリ＋アクティブなタスクブックパス」**のような概念をアプリ側で明示管理する必要がある。

3. **インストール単位**  
   Python ランタイム・依存パッケージ・（必要なら）Excel／xlwings の bundling 方針（インストーラ、社内配布用 zip、バージョン固定）を早期に決める。

4. **複数 Power Query パイプライン**  
   上記のとおり **UNC・「最新」指標・変換内容がクエリごとに異なる**。さらに **同一 UNC でも**「最新 1 ファイルだけ」と「新しい順に N 本を縦結合」など **ファイル集合のポリシーが異なる**。JavaFX 化では **ETL ジョブをクエリ単位（またはドメイン単位）にモジュール化**し、共有の「フォルダからファイル列を得る」ユーティリティを **日付列の意味（created / accessed / modified）・先頭 N 件・非表示除外**までパラメータ化すると再発バグを減らせる。

5. **Excel 名前定義に依存する PQ**  
   `_q結果_配台表` のように **`Excel.CurrentWorkbook()` で名前を引く**クエリは、JavaFX では **名前定義を廃し**、**出力ルートパスをアプリ設定で明示**する。読み取りは **PQ なしでファイル API 十分**なことが多い。

6. **選択的 Java 化の境界**  
   UI を JavaFX にした後も **計算の大半は Python のまま**が前提。**Java に移す候補**は、(a) **プロファイルで全体時間の有意な割合**を占める処理、(b) **入力出力がファイル／行列として明確**でゴールデンテストしやすい処理、(c) **業務ルールの変更頻度が低い**純粋計算、などに限定する。**Gemini 連携・環境依存が強い部分**は Python 継続が現実的なことが多い。

## 推奨フェーズ（ロードマップ）

### フェーズ 0: 棚卸しと非機能要件の固定（短サイクル）

- VBA から呼ばれる **全エントリ**（段階1/2、ガント更新、配台試行順、パターン別段階2 等）を一覧化。根拠: [code/python/xlwings_console_runner.py](code/python/xlwings_console_runner.py) のエントリ一覧と [planning_core/__init__.py](code/python/planning_core/__init__.py) の公開 API 説明。
- 各機能が **cmd 起動で完結するか / xlwings 必須か**をマトリクス化（移行難易度の見積りに直結）。
- **[plan/UI参照用_生産管理_AI配台(RC1).xlsx](plan/UI参照用_生産管理_AI配台(RC1).xlsx)** から **画面・シート・ボタン対応表**のドラフトを作る。**グラフ・チャート**が業務上どこまで必要か（既存ブック内のグラフ対象・種類）を **棚卸し**し、**JFreeChart 優先／JavaFX Chart 選択**の適用範囲の入力とする。
- **Power Query 依存**を洗い出し（現時点: **加工計画DATA（単一ファイル）**、**_q加工計画DATA_実績比較用（同一フォルダ・複数ファイル結合）**、**_q加工実績明細DATA**、**_q結果_配台表（名前「フォルダパス」＋ `結果_配台表.xlsx`）**）。共有フォルダ系については「PQ 代替」のどれを主経路にするか **意思決定**し、**ゴールデンファイル比較**の検証方針を書く。実績明細系は **累積実績・累積完了率・停機加算日時**を含め **行単位・集計列の双方**を検証する。実績比較用は **複数 `抽出時間` の縦結合後の行数・並び**に加え、**Upsert コメントと最終 M の差分**を仕様として確定する。**結果_配台表**は **出力ファイル直読＋列型マニフェスト**で十分か確認し、名前定義の移行先を決める。

### フェーズ 1: JavaFX プロジェクト基盤（新規）

- **`code_java/`** を Java プロジェクトのルートとし、**Maven ＋ JavaFX SDK（モジュールパスまたは依存）**を追加する（**`code/` 配下には Java プロジェクトを置かない**）。依存には **ControlsFX**（SpreadsheetView）・必要に応じ **Apache POI**・**JFreeChart** を **バージョン固定で**宣言する（**JavaFX Chart** は JDK/JavaFX に同梱のため依存追加は原則不要。使用画面のみ import）。**各依存は「OSS の永続無料」方針に適合していることを選定時に確認**し、TODO **`oss-perpetual-free-audit`** で記録する。
- パッケージング方針: `jlink` / `jpackage`（Windows 向け .exe 配布を想定する場合）。

### フェーズ 2: 「シェルアプリ」— 計算は Python、そのまま（MVP）

- JavaFX から `ProcessBuilder` 等で **`py -3.x` + 既存スクリプト**（例: [code/python/plan_simulation_stage2.py](code/python/plan_simulation_stage2.py)、[code/python/task_extract_stage1.py](code/python/task_extract_stage1.py)）を実行。
- 実行前に **`TASK_INPUT_WORKBOOK`** および必要なら **`PYTHONUTF8` 等**を子プロセス環境に設定（VBA がしていることの代替）。
- ログは **stdout/stderr または [code/python](code/python) 既存の log 配下**を UI でテール表示。終了コードは VBA が参照している `log/stage_vba_exitcode.txt` 等と整合させるか、Java 側はプロセスの exit code を主に見る設計に整理。

この段階で **「Excel を開かずに配台を回せる」**最低限の価値が出る（データ編集はまだ Excel でも可）。

### フェーズ 3: データ編集 UI の置換（コア工数）

- 主要シート（例: 「加工計画DATA」「配台計画_タスク入力」）を **OSS のスプレッドシート系（既定案: ControlsFX SpreadsheetView）を優先**して編集し、単純フォーム部分は **TableView／フォーム**と組み合わせる。保存時に **openpyxl と列互換のある xlsx/xlsm**として書き出す（**Apache POI** での検証・往復を検討）。
- **加工計画DATA**、**加工計画DATA_実績比較用（_q 相当）**、**加工実績明細系** について、編集 UI と別に **「共有フォルダから再取得（PQ 相当）」**アクションを JavaFX に用意する想定（フェーズ 0 で決めた ETL 経路に接続。クエリが複数あれば **取得ボタンを分けるか一括ジョブにするか**を UI 設計で決める。同一フォルダでも **単一版と N ファイル結合版は別ジョブ**）。
- **結果_配台表（`_q結果_配台表` 相当の画面）** は、配台実行後に生成される **`結果_配台表.xlsx`** を **出力フォルダから直接読み込み**して表示する（PQ 再現より **読込 API の方が優先**）。リフレッシュは「最新の配台結果を再読込」で足りる。
- 列定義の正は Python 側の定数（`_core.py` の `PLAN_*` / `TASK_*` 等）に合わせる必要があるため、**単一の「列定義マニフェスト」（JSON または共有生成スクリプト）**を検討すると保守が楽。
- **グラフ・ダッシュボード**: フェーズ 0 の棚卸しに基づき、**JFreeChart を既定**として実装し、**JavaFX Chart** は要件・レイアウトに応じて **選択的に併用**。埋め込み方式（**SwingNode** vs **画像 `ImageView`**）と **更新頻度・パフォーマンス**は TODO **`charts-jfreechart-javafx`** で POC 確定。

### フェーズ 4: xlwings 必須機能の縮小または代替実装

- フェーズ 0 のマトリクスに従い、残った xlwings 依存を **Python 側でファイル I/O 化**するか、限定的な別手段を選ぶ。

### フェーズ 5: VBA／Excel 依存の縮退

- 運用で JavaFX が主流になったら、マクロブックは **互換エクスポート目的のみ**に縮小、または廃止を検討。

### フェーズ 6（任意）: 重い処理の Java 化検討

- **前提**: フェーズ 2〜4 が運用可能で、**実データでのボトルネックが判明**していること。
- **手順**: JVM と Python 双方で **同じ入力**に対する **時間・メモリの計測** → ホットスポットを **関数／モジュール粒度**でリスト化 → **`code_java/` 配下に Java モジュール**として実装（JNI、プロセス間通信、または CSV バイナリ経由のパイプなど）→ **ゴールデン一致テスト** → フラグで切替。
- **スコープ管理**: 「全部 Java」ではなく **検証可能な単位**まで。**配台ルールの文章の正**（[配台ルール.md](配台ルール.md)）と実装の両方を更新する必要がある変更は、Python/Java のどちらを正にするか事前に決める。

## ドキュメント・ルールとの関係

- **配台ロジックを変える場合**は [配台ルール.md](配台ルール.md)（リポジトリ内の業務文章の正）との整合が必要（`.cursor/rules/dispatch-docs-sync.mdc` の趣旨）。
- **巨大ファイル** `_core.py` の読み方は [code/python/planning_core/_core_FILE_MAP.txt](code/python/planning_core/_core_FILE_MAP.txt) を前提に局所調査する（`.cursor/rules/planning-core-huge-file.mdc`）。

## 実装前に固める取り決め（決定値）

本節は **「§ 意思決定が必要な論点」の決定結果**を確定値として記録する。**未決事項は明示的に「保留」と書き、推測で埋めない**。記載が下表の「§ 意思決定が必要な論点」と重複するときは **本節を正**として扱う。

### L0. 基盤・ツールチェーン（決定）

- **OS**: **Windows 11 限定**（Windows 10 は対象外）。
- **JDK / JavaFX**: **最新 LTS** を採用（Temurin 系 / OpenJFX 系の同 LTS 系列。**具体バージョン番号は `decision-l0-toolchain-finalize` で固定**）。
- **ビルドツール**: **Maven**（`code_java/` 直下に `pom.xml`、必要に応じてマルチモジュール）。
- **配置**: **`code_java/`** に集約（既存 `code/` には Java ソース・ビルド成果物を置かない）。
- **JVM 既定**: `-Dfile.encoding=UTF-8`、`Locale.JAPAN`、`ZoneId.of("Asia/Tokyo")`。

### Excel との論理的切り離し・互換出力・致命的エラー（決定）

- **ロジックの正**: **JavaFX＋Python（`planning_core` 等）のみ**。Excel は **入出力ファイルの読み書き先としての形式**にはなるが、**実行時の条件分岐・フォールバック先としては用いない**。
- **Excel 可読な成果物**: **`.xlsx`（および必要に応じ CSV 等）を最終成果の標準**とし、Excel で開けることを受け入れ基準に含める。
- **フォールバックの範囲**: **同一プロセス／パイプライン内**（例: 入力検証の再試行、部分的な設定修正の促し、安全な縮退モードが定義されている場合のみ）。**Excel マクロや手作業の値貼りを前提とした迂回は定義しない**。
- **致命的エラー**: データ整合が取れない・リソース不在で継続不能・ブック破損リスクがある等、**誤った結果を出し続けるより中断が適切**な場合は **処理を停止**する。
- **ユーザー通知**: 致命的エラーおよび **`PlanningValidationError` 等の業務エラー**では、**原因（識別可能な説明）**と **推奨対処法**を **同一ダイアログまたはログの冒頭**に必ず含める（「不明なエラー」だけで終わらせないことを目標とする）。

### 1. Python ↔ JavaFX の連携契約

| 項目 | 決定 |
|------|------|
| **1.1 Python 実行系** | システムの `py -3.x` を使用。`code_java`（または展開先）配下に **専用 venv** を **初回起動バッチで構築**（環境汚染防止＋手軽さの両立）。 |
| **1.2 パッケージ更新方針** | **`requirements.txt` で完全固定**（`pip freeze` 相当）。再現性を最優先。 |
| **1.3 IPC プロトコル** | **引数 ＋ 環境変数 ＋ stdout への NDJSON**。進捗イベントは構造化し、JavaFX 側で ProgressBar 等にバインド可能にする。 |
| **1.4 進捗・ログのストリーム** | **stdout = 進捗（NDJSON）**／**stderr = ログ**。JavaFX 側でキャプチャし、UI 更新と `logs/` への追記を両立。 |
| **1.5 キャンセル／中断** | **フラグファイル方式**（`.cancel` ファイル生成）。Python 側ループ境界で存在チェックし安全終了。**強制 kill はしない**（ブック破損回避）。 |
| **1.6 終了コード体系** | `0`=正常／`1`=予期せぬエラー／`2`=**致命的エラー**（フォールバック不可・処理継続不能・**原因と推奨対処を stderr／NDJSON で必須**）／`3`=`PlanningValidationError`（既存・データ検証）／`9`=ユーザー中断。 |
| **1.7 エラー伝播・例外の翻訳** | **`3` および `2`**: Python 側で **stderr または NDJSON** に **理由・推奨対処法**を構造化して出力 → JavaFX で **日本語ダイアログ**に落とし込む。**`1`**: 可能な範囲でスタック要約とログ参照先を表示。**「不明なエラー」だけで終わらせない**ことを原則とする。 |
| **1.8 同時実行ポリシー** | **UI レベルで完全な単一実行**。実行中は実行ボタンを無効化し、多重起動による書き込み競合を防ぐ。 |

#### IPC テンプレート（ドラフト・AI 生成コードの固定用）

本項は **フィールド名・構造を実装前に固定するためのドラフト**である。リポジトリに **`ipc-line.schema.json`**（stdout 1 行分の JSON Schema）および **`columns.schema.json`**（`schema.json` インスタンス用）を配置する場合は、本節を **正本**として同期する。**実装済み**: [code_java/src/main/resources/schema/ipc-line.schema.json](code_java/src/main/resources/schema/ipc-line.schema.json)、[code_java/src/main/resources/schema/columns.schema.json](code_java/src/main/resources/schema/columns.schema.json)、検証用インスタンス例 [code_java/src/main/resources/schema/example-column-manifest.json](code_java/src/main/resources/schema/example-column-manifest.json)。スキーマ本文の `description` は UTF-8 互換のため英語のみ。

**規約（stdout）**

- **1 行 1 JSON**（NDJSON）。改行は `\n` のみ。ストリーム途中で **途中分割された JSON は書かない**（バッファリングしてから 1 行書き出す）。
- **文字コード UTF-8**。先頭 BOM は付けない。
- **必須共通フィールド**（全 `type`）: `schemaVersion`（文字列、当面 `"1.0"`）、`type`（列挙）、`ts`（ISO-8601、タイムゾーン付き推奨、例 `2026-05-02T13:00:00+09:00`）。

**`type` 別フィールド（ドラフト）**

| `type` | 用途 | 追加フィールド（推奨名・型） |
|--------|------|------------------------------|
| `progress` | 進捗・ProgressBar | `step`（string）、`label`（string・画面表示用）、`current`（int）、`total`（int・省略可）、`percent`（number 0〜100・省略可） |
| `log` | 構造化ログ（UI のログペインへも） | `level`（`DEBUG` \| `INFO` \| `WARN` \| `ERROR`）、`message`（string）、`logger`（string・省略可） |
| `validation_error` | データ検証・業務エラー（終了コード **3** 想定） | `exitCode`（固定 **3**）、`code`（string・機械可読）、`message`（string・ユーザー向け要約）、`remediation`（string の配列・推奨対処を順に）、`detail`（object・省略可） |
| `fatal_error` | 致命的エラー・継続不能（終了コード **2** 想定） | `exitCode`（固定 **2**）、`code`、`message`、`remediation`（配列）、`cause`（string・技術的理由・省略可）、`detail`（object・省略可） |
| `done` | 正常終了直前（終了コード **0** とセット） | `exitCode`（固定 **0**）、`outputPaths`（string の配列・生成した成果物パス）、`summary`（string・省略可） |
| `ping` | 長時間処理の生存確認（省略可） | `label`（string・省略可） |

**規約（stderr）**

- **人間可読・ライブラリ既定ログ**を想定し、**プレーンテキスト行**を許容する（構造化パースは **必須としない**）。
- **ユーザー向けの原因・推奨対処は stdout の `validation_error` / `fatal_error` に必ず出す**（stderr のみに頼らない）。
- 例外スタックトレースは stderr にそのまま出してよいが、JavaFX は **`fatal_error` NDJSON を優先**してダイアログ文言を組み立てる。

**stdout の最低限サンプル（NDJSON・抜粋）**

```json
{"schemaVersion":"1.0","type":"progress","ts":"2026-05-02T13:00:01+09:00","step":"stage1","label":"段階1 抽出","current":1,"total":4,"percent":25}
{"schemaVersion":"1.0","type":"log","ts":"2026-05-02T13:00:02+09:00","level":"INFO","message":"入力ブックを読み込みました","logger":"planning_core"}
{"schemaVersion":"1.0","type":"validation_error","ts":"2026-05-02T13:00:10+09:00","exitCode":3,"code":"MISSING_COLUMN","message":"必須列が不足しています","remediation":["schema.json と入力ファイルの列名を照合してください","加工計画DATA の取込をやり直してください"],"detail":{"sheet":"加工計画DATA","column":"依頼NO"}}
{"schemaVersion":"1.0","type":"fatal_error","ts":"2026-05-02T13:00:11+09:00","exitCode":2,"code":"UNC_NOT_REACHABLE","message":"共有フォルダに接続できません","remediation":["VPN／ネットワークを確認してください","settings.json のパス設定を確認してください"],"cause":"java.nio.file.FileSystemException: ..."}
{"schemaVersion":"1.0","type":"done","ts":"2026-05-02T13:05:00+09:00","exitCode":0,"outputPaths":["output/production_plan_20260502_130500.xlsx"],"summary":"段階2 完了"}
```

**列定義マニフェスト `schema.json`（ドラフト・構造）**

- **配置例**: `code_java/src/main/resources/schema/manifest.schema.json`（JSON Schema）と **`schema.json` 実体**（リポジトリ共通パスは `decision-2-data-contract` で最終決定）。
- **目的**: シート単位で **列ヘッダー・型・Python 定数名・Java プロパティ名**を単一ソース化する。

**インスタンス例（ドラフト・一部のみ）**

```json
{
  "schemaVersion": "1.0",
  "document": "dispatch-plan-column-manifest",
  "sheets": [
    {
      "id": "machining_plan_data",
      "excelSheetName": "加工計画DATA",
      "columns": [
        {
          "headerKey": "依頼NO",
          "pythonConstantHint": "TASK_COL_TASK_ID",
          "javaProperty": "taskId",
          "valueKind": "string",
          "nullable": false
        },
        {
          "headerKey": "換算数量",
          "pythonConstantHint": "TASK_COL_QTY",
          "javaProperty": "convertedQty",
          "valueKind": "decimal",
          "nullable": true
        }
      ]
    }
  ]
}
```

**フィールド説明（ドラフト）**

| フィールド | 説明 |
|------------|------|
| `schemaVersion` | マニフェスト形式の版（`1.0` から開始）。 |
| `sheets[].id` | 安定 ID（コード・設定から参照）。 |
| `sheets[].excelSheetName` | Excel シート名（openpyxl／POI での一致用）。 |
| `columns[].headerKey` | シート上の列見出し文字列（**Python `_core.py` の定数と一致**させる）。 |
| `columns[].pythonConstantHint` | 対応する `PLAN_*` / `TASK_*` 名（ドキュメント・検証用、実装では `_core.py` が真）。 |
| `columns[].javaProperty` | JavaFX／POI 側でのプロパティ名（キャメルケース推奨）。 |
| `columns[].valueKind` | `string` \| `integer` \| `decimal` \| `date` \| `datetime` \| `boolean` 等（境界での変換ルールの前提）。 |
| `columns[].nullable` | null 許容か。 |

### 2. データコントラクト（列定義／ファイル）

| 項目 | 決定 |
|------|------|
| **2.1 列定義マニフェスト** | **`schema.json` を正**として `code_java/` 等の共通領域に配置。**Python と JavaFX の双方が起動時に読み込んで列構成を認識**する。**フィールド名・構造のドラフトは §「IPC テンプレート（ドラフト）」の `schema.json` 例を参照**。正式版は JSON Schema（`columns.schema.json`）で検証可能にする。 |
| **2.2 「設定_環境変数」シートの後継** | アプリ実行ディレクトリ直下の **`settings.json`**。UI に設定画面を設け、更新時に上書き保存。 |
| **2.3 マクロブック（.xlsm）の取り扱い** | UI 置換後は **POI／openpyxl 等のファイル API で読み書きを完結**。**xlwings／Excel 自動操作は撤廃**（ロジックが Excel ランタイムに依存しないことを検証する）。 |
| **2.4 入出力ファイルの命名・配置** | 例: `output/production_plan_yyyyMMdd_HHmmss.xlsx`。**出力先フォルダ固定 ＋ 日時サフィックス**を必須化。 |
| **2.5 結果_配台表 出力ルート** | `settings.json` に **`RESULT_OUTPUT_DIR`** キーを設ける。**既定は相対 `./output`**、絶対パスへ変更可能。**旧 PQ の名前定義「フォルダパス」は廃止**。 |
| **2.6 タイムゾーン／日付シリアル変換** | JavaFX 内部はすべて **`LocalDate` / `LocalDateTime` で統一**。**変換は POI 入出力境界のみ**で行う（Excel シリアル値・書式は境界で吸収）。 |

### 3. PQ 代替 ETL（PQ-A〜D の Python 同等実装）

| 項目 | 決定 |
|------|------|
| **3.1 採用方針** | **「Python ETL」で確定**。JavaFX の取得実行ボタン → Python スクリプトを呼び出してデータ生成。 |
| **3.2 共通ユーティリティ API** | **`fetch_latest_files(unc_path, date_kind='created', top_n=1)`** を共通関数として定義し、PQ 代替処理で再利用。 |
| **3.3 出力フォーマット** | 当面は **xlsx**（後段 `planning_core` の互換維持）。将来的な高速化が必要になれば Parquet 等を検討。 |
| **3.4 ゴールデンデータ** | 匿名化したテスト用 xlsx を **`tests/fixtures/`** に保管し、回帰テスト（pytest）の正本とする。 |
| **3.5 PQ-B の Upsert 挙動** | **同一キー（依頼NO 等）が複数あれば、`抽出時間` 最新の行を残す（Drop Duplicates）**仕様で実装する。 |
| **3.6 `ファイルの変換 (2)/(3)` の実体** | **UI 参照ブックから抽出した M を一次資料**に、pandas の型変換・列選択ロジックとして実装に落とし込む（TODO: `pq-m-export-from-workbook`）。 |

### 4. 画面（IA・遷移・状態管理）

| 項目 | 決定 |
|------|------|
| **4.1 トップレベル IA** | **左サイドバー（機能メニュー）＋ メインエリア（タブ切替）**のモダン構成。 |
| **4.2 主要画面（暫定）** | **「設定」／「実績データ取込」／「計画データ編集（SpreadsheetView）」／「結果設備ガント」**の **4 主要画面**に整理。 |
| **4.3 ワークスペース概念** | **導入しない**。常に **`settings.json` で指定した「最新の入力ファイル」のパス**を直接読み書きする **単一運用**。 |
| **4.4 状態遷移** | **編集（未保存 `*` マーク）→ 保存（必須）→ 実行ボタン有効化 → 実行中（プログレス表示）→ 完了（結果画面へ自動遷移・再読込）**の **単方向サイクル**。 |
| **4.5 キーボード操作・IME・コピペ** | ControlsFX `SpreadsheetView` の標準機能を採用。**Excel ↔ JavaFX のクリップボード往復**で不足する部分のみ API 実装で補う。 |

### 5. 配布・運用・セキュリティ

| 項目 | 決定 |
|------|------|
| **5.1 インストーラ／配布形式** | **社内配布用 ZIP**（解凍 → `run.bat` で起動）。インストーラの権限問題を回避し、デプロイの手軽さを優先。 |
| **5.2 Python ランタイムの同梱／更新** | アプリ更新と分離。**初回起動バッチで `pip install -r requirements.txt`** を実行して環境を揃える。 |
| **5.3 自動更新／バージョニング** | **日付ベース**（例: `v2026.05.02`）。共有フォルダに最新 ZIP を置き、**手動で差し替える運用**。**後方互換は原則 1 世代前まで**担保。 |
| **5.4 資格情報（API キー等）の管理** | **配布対象 ZIP・設定には含めない**。OS の **ユーザー環境変数**または **配布対象外のローカル `.env`** ファイルで個別に管理する。 |
| **5.5 UNC 共有フォルダ障害時の挙動** | 読込・書込の直前に **`Path.exists` 等でアクセスチェック**。失敗時は **即「ネットワークエラー」ダイアログ**を表示（**自動リトライは無限ループ回避のため行わない**）。 |
| **5.6 ログの場所と保持期間** | アプリフォルダ直下 `logs/` に **日別ローテーション**で出力。**保持期間 14 日**。**機微情報はマスキング**して出力。 |

### 6. 品質保証・テスト戦略

| 項目 | 決定 |
|------|------|
| **6.1 テスト方針** | **UI = 手動 QA**、**ETL／配台コア = pytest によるゴールデン比較**。 |
| **6.2 ゴールデン許容誤差** | **日付・文字列・整数（数量等）は完全一致**。**浮動小数（完了率等）は小数第 2 位までの丸め一致**を許容条件とする。 |
| **6.3 CI／自動化** | 当面ローカル環境想定。**`test.bat`（Maven test ＋ pytest の一括走行）**を CI 代替とする。 |
| **6.4 TestFX** | UI が手動 QA のため **依存に含めない**（OSS 永続無料監査の対象外）。 |

### 7. ドキュメント・正本管理

| 項目 | 決定 |
|------|------|
| **7.1 プランファイルの同期手順** | Cursor の `.cursor/plans` のファイルを **手作業で**リポジトリ `plan/` フォルダへ上書きコミット（**フック自動化はしない**＝仕組みの複雑化を避ける）。 |
| **7.2 配台ルール.md の正の決め** | ロジックの正は引き続き **Python（`_core.py`）と `配台ルール.md`** に置く。**JavaFX は「データの受け渡しと表示の器」**に限定。 |
| **7.3 OSS 永続無料台帳のフォーマット** | **`plan/oss_licenses.md`**（Markdown 表形式：**コンポーネント名／バージョン／ライセンス種別／利用可否**）で管理する。 |
| **7.4 ライセンス通知（NOTICE）** | プロジェクトルートに **`NOTICE.txt`** を作成し、**配布用 ZIP にも同梱**する。 |

### 8. 失われやすい運用論点

| 項目 | 決定 |
|------|------|
| **8.1 移行検証と二重運用** | **業務上の正は JavaFX＋Python ライン**とする。移行検証期間中は **旧 Excel 版の出力と新アプリの出力を比較**して差分を確認するが、**「Excel UI が正」の二重運用は行わない**。同一データの同時編集は避ける（ファイルロック・単一実行ポリシーと整合）。 |
| **8.2 マスタ（基礎データ）** | **基礎データは Excel ブックで提供する**（`.xlsx` / `.xlsm`）。**編集・メンテナンスは Excel を主とする**運用とする。**アプリ・Python は openpyxl／POI のファイル読込のみ**とし、マクロ実行や Excel 常駐を前提にしない。リポジトリ例: [plan/master.xlsx](plan/master.xlsx)。**将来的に JavaFX からマスタを編集する UI を付けるかは別判断**（工数・優先度）。 |
| **8.3 フォールバック・致命的エラー** | **フォールバックはロジック内で完結**（§「Excel との論理的切り離し…」参照）。**フォールバック不可能な場合は終了コード `2` 等で致命的エラーとして中断**し、**原因と推奨対処法を UI で明示**する。**旧 Excel への値貼り付け・VBA 実行による迂回は採用しない**。 |
| **8.4 配台パターンサマリの UI 仕様** | JavaFX 画面上に **チェックボックス／ラジオボタン等の簡易フォーム**を設け、**選択値を Python のプロセス実行引数へ直接渡す**仕様とする。 |

## 意思決定が必要な論点（実装前に固めると安全）

> **更新**: L0 と 1〜8 の決定値は **§ 実装前に固める取り決め（決定値）** に確定済み。下表は **論点の歴史記録**として残す。最新の正は上記節を参照。

| 論点 | 選択肢のイメージ |
|------|------------------|
| 計算コア | **既定: Python 継続**。重い部分のみ **計測後に Java 部分移植（ハイブリッド）**。**全面 Java 移植はスコープ外**として扱う |
| **Java ソースのルート** | **`code_java/`** に統一。**`code/` は既存 Excel／Python ラインのまま**とし衝突を避ける |
| **OSS の永続無料** | **必須**。有償サブスク・オープンコアの実質有償・二重ライセンスで有償が必要な解釈は原則採用しない。**依存台帳（`oss-perpetual-free-audit`）**で証跡を残す |
| **グリッド UI（Excel 親和・OSS）** | **既定: Java と親和性の高い OSS を優先**（**ControlsFX SpreadsheetView** を POC 起点）。単純一覧は **TableView** で可。**Apache POI** で書式・往復検証を検討。**永続無料**・ライセンス・保守で最終確定 |
| **チャート／グラフ** | **既定: JFreeChart を優先**。軽量・シンプルな系列は **JavaFX Chart** を **状況に応じ選択**。埋め込み（SwingNode／画像）・LGPL 可否は POC・法務で確定 |
| 保存形式 | 当面 **xlsm 互換** vs 早めに **DB/JSON＋エクスポート** |
| xlwings 代替 | **Python をファイル中心に寄せる（確定）**。Excel 常駐（xlwings）は **不採用**（§「Excel との論理的切り離し」） |
| **加工計画DATA 供給** | **Python ETL で PQ 同等（推奨して検証）** vs バッチ vs Excel 自動化のみ |
| **加工計画DATA_実績比較用 供給** | 同上（**N ファイル縦結合・複数 `抽出時間`**。単一ファイル版との **共通化と差分**をコードで分離） |
| **加工実績明細（_q）供給** | 同上（**累積・完了率・複数日付キー**で検証コストが一段高い） |
| **結果_配台表の参照** | **出力 xlsx 直読＋設定で出力パス**（名前「フォルダパス」を廃止）。列型は M の `TransformColumnTypes` と整合 |

---

**まとめ**: **主目的は JavaFX UI**。**業務ロジックは Excel／VBA／COM から切り離し**、**成果物は Excel が開ける `.xlsx` 等で提供**する。**フォールバックはロジック内のみ**とし、**不可能なら致命的エラーで中断して原因・推奨対処を通知**する。**Java と親和性の高い OSS を積極利用**し、採用 OSS は **永続無料で利用できる条件のものに限定**する。**Excel に近いグリッド UX**（既定案は ControlsFX SpreadsheetView 起点の POC）と **グラフ・可視化**（**JFreeChart 優先**、**JavaFX Chart は選択的**）を目指す。新規 Java／JavaFX は **`code_java/`**、既存資産は **`code/`** と **ツリー分離**して衝突を避ける。Python は当面そのまま呼び出し、**全コードの Java 書き換えは求めない**。必要になれば **計測に基づき重い処理だけ Java 化**するオプションを残す。**複数 Power Query 由来データ**や **xlwings 撤廃／PQ 代替**が技術的な主な負荷であり、**UI 移行とデータパイプライン整理が先行**、高速化はその後の選択となる。**実装着手前に固める基盤・IPC・データ契約・ETL・画面・配布・QA・ドキュメント・運用の各取り決めは §「実装前に固める取り決め（決定値）」を正**とし、ブレずに進める。
