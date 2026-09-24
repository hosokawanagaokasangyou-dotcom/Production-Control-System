# 工程管理AIへ後加工工賃を統合する実装指示

日付: 2026-09-16  
対象リポジトリ: `C:\システム開発\工程管理AIプロジェクト_JAVA`  
この文書は自己完結。別チャットへこのファイルだけ渡せば実装できる。コードはこの文書の作成時点では未着手。

## 1. 目的

工程管理AIデスクトップ（JavaFX・同じ exe）に、後加工工賃の業務画面をメインタブとして追加する。

- 見た目・操作感は既存シェルに揃える（上部ツールバー、工場コンボ、Busy、ステータスバー、環境変数タブ、既存 RDP）
- 結果はアプリ内でも見る（KPI・表・グラフ・メールコピー）。Excel も現行スキルと同じ場所・名前で出す
- 利用者は Cursor も Python も独立 KouchinVerify exe も不要。工程管理AIを起動して「後加工工賃」タブを使う

## 2. やってはいけないこと

- 配台・依頼書転記・受注検索・マスタ編集・勤怠・ガントのロジックを工賃のために書き換えない
- 既存メインタブ「加工トレンド」（子タブ「加工量」「加工賃」）と混同しない
  - 既存: 受注ファイル AH（加工賃単価）× 日次の工程延べ m → 円の日次/累計。データは日報・明細・アラジン・配台
  - 今回: ①東レ検収CSV・②長岡明細（国分）/加工賃試算（湖南）・③アラジン依頼NO別問合せ の**月次金額突合**と、②からの**工場間・工程別月次トレンド**
- Python tkinter / PyInstaller（`●自動検証\app` / `kouchin-python-release`）を工程管理AIに埋め込まない
- 工程管理AI本体を工賃専用リポジトリへコピーして二重管理しない
- 工賃タブにリモートデスクトップ画面を複製しない（既存「リモートデスクトップ」タブを使う）
- `~/.kouchin-verify-desktop` を新規の正本にしない。永続化は `~/.pm-ai-desktop/session-state.json`

## 3. 移植元（正本）

| 役割 | 場所 |
|------|------|
| 検証コア（Java・大半は移植済み） | `C:\システム開発\後加工工賃検証_JAVA\code_java\src\main\java\jp\co\pm\kouchin\verify\` 入口は `VerifyService.java` |
| 検証の仕様・Python 正本 | `\\192.168.0.101\共有フォルダ\国分工場\国分管理\後加工\後加工工賃明細\●自動検証\.cursor\skills\kouchin-consistency-check\`（`SKILL.md` と `scripts\verify_kouchin.py`） |
| 月次トレンド（Python・Java 未移植） | 同上 `●自動検証\.cursor\skills\kouchin-monthly-trend\`（`SKILL.md` と `scripts\`） |
| 月次トレンド設計 | `●自動検証\docs\specs\2026-09-07-kouchin-monthly-trend-design.md` |
| 独立 JavaFX シェル（参考のみ。コピーしない） | `C:\システム開発\後加工工賃検証_JAVA\code_java\src\main\java\jp\co\pm\kouchin\desktop\` |

Java 検証で Python より足りない可能性があるもの（実装時に `verify_kouchin.py` と突き合わせる）:

- 検証C（湖南・月次処理ファイルの東レ合計照合）
- 原本コピーシートのハイライト
- 手動判定.csv / 前月過不足.csv の扱いが Python と同等か

足りなければ Python 仕様に合わせて Java コアを補ってから UI に載せる。

## 4. パッケージ配置（工程管理AI側）

推奨（これで固定する）:

- UI: `jp.co.pm.ai.desktop.kouchin`
- 検証ロジック: `jp.co.pm.ai.kouchin.verify`（独立リポジトリの `jp.co.pm.kouchin.verify` を**ソースコピーしてリネーム**）
- 月次トレンドロジック: `jp.co.pm.ai.kouchin.trend`（Python を新規移植）
- FXML: `code_java/src/main/resources/jp/co/pm/ai/desktop/fxml/` 例:
  - `KouchinHostTab.fxml`（子タブホスト）
  - `KouchinVerifyTab.fxml`
  - `KouchinTrendTab.fxml`
  - 任意 `KouchinMailTab.fxml`
- テスト: `code_java/src/test/java/jp/co/pm/ai/desktop/kouchin/` と `jp/co/pm/ai/kouchin/verify/` / `trend/`

独立 exe（`後加工工賃検証_JAVA` の jpackage、`kouchin-python-release`）はメンテしない。工程管理AIの `package-portable` で配る。

Maven の multimodule 化はしない（既存単一 `code_java` にソースを入れる）。

## 5. メインタブ登録（必須チェックリスト）

ルール正本: `C:\システム開発\工程管理AIプロジェクト_JAVA\.cursor\rules\main-shell-tab-management.mdc`

新規メインタブ:

- enum: `MainShellTabId.KOUCHIN`、key `"kouchin"`
- 見出し: 「後加工工賃」
- 配置: トップレベル。`MainShellTabLayoutDefaults.DEFAULT_FLAT_TAB_KEY_ORDER` の末尾（タブ整理の直前）に `"kouchin"` を明示追加
- `groupedLayout()` にも入れる（仕様が無ければトップレベル末尾）

子タブ（ホスト FXML の `Tab` の `text` と一字一句一致）:

1. 検証
2. 月次トレンド
3. 報告メール（任意。検証タブ内ペインでも可。子タブにするなら catalog に含める）

必ず更新するファイル:

1. `MainShellTabId.java`
2. `MainShellTabLayoutDefaults.java`（DEFAULT と groupedLayout）
3. `MainShell.fxml`（`Tab` + `fx:include`。他タブと同様 **lazy include**。未選択時に xlsm/CSV を読まない）
4. `MainShellController.java`（`@FXML`、`mainShellTabId` / `mainShellTabFor`、選択時フック。加工トレンドの `onMainShellTabSelected` と同様）
5. `MainShellInnerTabCatalog.labelsFor` に `KOUCHIN -> List.of("検証", "月次トレンド")`（メールを子タブにするなら3つ目を追加）
6. FXML ロードテスト（`ProcessingTrendHostTabFxmlTest` / `ProcessingFeeTrendTabFxmlTest` と同型。見出し文字列を assert）

`MainShellTabId` だけ足して Defaults / InnerTabCatalog を忘れるのは禁止。

## 6. 環境変数・パス

ルール正本: `.cursor/rules/env-vars-managed-by-sheet-and-tsv.mdc`

- 実行時の正本は環境変数タブ → `~/.pm-ai-desktop/session-state.json` の `uiEnvRows`
- キー追加時は `AppPaths`、`EnvVarDocs`、`ui_ref_env_defaults.json` を同時更新
- 工場コンボ（`shellFactorySiteCombo`）は国分/湖南の**既定フォルダ切替**に使う。検証の「両工場まとめて」は**別ボタン**（コンボを国分にしただけで湖南が走らないようにする）

既定 UNC（スキルと同じ。変更しない）:

- 出力・①CSV・国分③・手動判定: `\\192.168.0.101\共有フォルダ\国分工場\国分管理\後加工\後加工工賃明細\●自動検証`
- ①: その下 `東レ送付CSV\RVSHEETyyyymm.csv`（入庫場所 A010=国分、A010P=湖南）
- 国分②: `●自動検証\国分工場\長岡後加工賃明細`（検証）。月次トレンドの国分②は年度フォルダ直参照:
  - `\\192.168.0.101\共有フォルダ\国分工場\国分管理\後加工\後加工工賃明細\工賃明細2026年度（令和8年度）`
  - `\\192.168.0.101\共有フォルダ\国分工場\国分管理\後加工\後加工工賃明細\工賃明細２０２５年度　(R7年度)`
- 湖南②: `\\192.168.0.101\共有フォルダ\湖南工場\湖南共有\002  加工G\000  後加工業務\2 後加工試算`（空白は半角2つ）。`.lnk` は読まない
- 湖南③: `\\192.168.0.101\共有フォルダ\湖南工場\湖南共有\生産管理システム\管理システム\●DATA\月次実績`
- 湖南月次処理（検証C）: `\\192.168.0.101\共有フォルダ\湖南工場\湖南共有\生産管理システム\アラジンオフィスシステムデータ\月次実績表\0 東レ月次処理ファイル`

環境変数キー例（実装で名前を決めたら EnvVarDocs に説明を書く）:

- 自動検証ルート（`--base` 相当）
- 湖南試算ルート / 湖南アラジン / 湖南月次処理ルート（上の既定）
- 国分明細年度フォルダ（トレンド用、複数可）
- 出力先（既定は自動検証ルート）

`.lnk` は使わない。絶対パスのみ。

## 7. シェルへの載せ方

- Busy: メインシェルの `shellStageProgressBox` / 既存 Stage 実行の中断ボタンに乗せる。工賃専用の別ウィンドウ Busy を常時出さない
- ログ: 検証子タブ内のログ欄＋下部ステータスバー
- RDP: 既存タブ。工賃から「RDPタブへ移動」リンク程度は可
- 文字コード: `.cursor/rules/java-utf8-string-literals.mdc`（UTF-8）
- ビルド: `.cursor/rules/code-java-maven-build.mdc`
- コミット時 version.txt: `.cursor/rules/version-txt-bump-on-commit.mdc`
- 依頼書原本フォルダは読み取り専用のまま（工賃はそこを書かない）

## 8. 検証機能（仕様の要約）

詳細は `kouchin-consistency-check/SKILL.md`。実装は Java `VerifyService` を UI から呼ぶ。

| 検証 | 工場 | 内容 |
|------|------|------|
| A | 両工場 | ① vs ② を契約NO（ハイフン除去） |
| B | 両工場 | ② vs ③ を依頼NO（枝番統合・月ずれ累計） |
| C | 湖南のみ | 月次処理ファイル東レ合計 vs 本検証総額 |
| D | 国分のみ | ②「東レまとめ」と元4シートの参照・式・金額 |

実行ボタン: 国分 / 湖南 / まとめて（国分→湖南、統合メール）。許容差既定 0.5 円。

出力（自動検証ルート）:

- `検証結果_国分工場_yyyyMMdd_HHmmss.xlsx`（+ html があれば同様）
- `検証結果_湖南工場_yyyyMMdd_HHmmss.xlsx`
- まとめて時 `報告メール_統合_yyyyMMdd_HHmmss.txt`
- 最新以外は `過去検証結果\` へ退避（Python と同じ）

対象月は①ファイル名 `RVSHEETyyyymm` の最大月。②③はそれに合わせる。過去月②は直近6か月。

## 9. 検証子タブ UI

実行前:

- 検出した①②③（と湖南の月次処理ファイル）のパス・対象月を表で見せる。無いものは赤字

実行:

- 国分 / 湖南 / まとめて。実行中はボタン無効、Busy＋中断

実行後:

- KPI カード: 検証A要確認、検証B要確認、報告する過不足、警告件数、（国分）検証D要修正、（湖南）検証C判定
- 警告リスト（①小計検算不一致、検証D など。先頭に出す）
- TableView: 検証A（契約NO）と検証B（依頼NO）。判定で色分け、オートフィルタ相当（判定・検索）。列は Excel に合わせる（契約NO/依頼NO/①②③金額/差額/報告計上額/判定/備考）
- 原本全行の TableView 化はしない。要確認＋検索のみ。詳細は「Excelを開く」
- 報告メール本文（国分→湖南の統合）。クリップボードコピー。txt/html 保存は既存仕様どおり
- 「Excelを開く」「出力フォルダを開く」

判定色の意味は SKILL.md「レポートの読み方」に従う（不一致・翌月記載・片側のみ・前月調整・前月過不足・手動判定・月ずれ解消・枝番統合）。

## 10. 月次トレンド機能

①③突合はしない。②のみ。Python `trend_kouchin.py` を Java 移植。

- 期間: 直近 N か月（既定 6）
- 国分: 年度フォルダの `*後加工工賃明細*.xls*`、年月はファイル名
- 湖南: 試算ルートの作業中 `月度加工賃試算.xlsm` と `yyyy年度試算　湖南\m月度加工賃試算.xlsm`
- 単位: 量 m→km、賃 円→千円（整数）
- 工程: 国分「スライス1/3」は「スライス」に合算して湖南と比較
- 全工程スパゲッティ折れ線は出さない

出力: `●自動検証\月トレンド_yyyyMMdd_HHmmss.xlsx`、旧ファイルは `過去月トレンド\` へ。シートはサマリ / 工程比較 / 負荷シフト / 付録。

UI:

- 月数、工場（国分/湖南/両方）
- サマリ: 合計・前月差・振替疑い Top（断定しない）
- 工程比較: 注目工程ごと 国分 vs 湖南の LineChart（量・賃）。`ProcessingTrendChartSupport` は描画参考にしてよいが、集計データは混ぜない
- 「Excelを開く」

## 11. 最適化

- 検証・トレンドはワーカースレッド。UI は `Platform.runLater` のみ
- UNC のファイル一覧は mtime キャッシュ。再実行/再検出で無効化
- メインタブ・子タブは lazy。開くまで POI/CSV しない
- 巨大原本の全セルを UI に載せない
- Prism / GPU プローブは既存 `PmAiFxApp` に任せる。工賃タブだけで `prism.order` を変えない

## 12. テスト

- `後加工工賃検証_JAVA` の JUnit を工程管理AIの `src/test` へ移し、パッケージ変更後も通す
- トレンド: Python `kouchin-monthly-trend/tests` 相当（読取・集計・Excel 最低限）を Java で書く。TDD（失敗するテスト→実装）
- `KouchinHostTabFxmlTest` 等で FXML がコントローラ付きでロードでき、子タブ見出しが catalog と一致
- 既存 `ProcessingTrendHostTabFxmlTest` など配台・加工トレンドのテストが落ちないこと
- UNC 実データでの両工場実行は手動スモーク。CI 必須にしない

## 13. 実装順序（推奨）

1. `jp.co.pm.kouchin.verify` を工程管理AIへコピーしパッケージ改名。テストが緑
2. Python との差分（検証C 等）を埋める
3. 空のメインタブ＋子タブ FXML をシェルに登録（チェックリスト消化、FXML テスト）
4. 検証 UI を `VerifyService` に接続（KPI・表・メール・Excel）
5. 月次トレンドを Java 移植＋チャート＋Excel
6. 環境変数キー・工場コンボ・Busy・lazy
7. 工程管理AIを起動して国分・湖南・まとめて・トレンドを確認。既存タブが壊れていないこと
8. 既存の portable 手順でパッケージ

## 14. 完了条件

- 工程管理AI起動 → メインタブ「後加工工賃」→ 検証（国分・湖南・まとめて）と月次トレンドが動く
- 結果 Excel / 統合メールが `●自動検証` に出る
- アプリ内で KPI・要確認表・（トレンド）グラフ・メールコピーができる
- 既存の配台・依頼書・「加工トレンド」（加工量/加工賃）が従来どおり動く
- `package-portable` 相当の成果物に工賃タブが含まれる

## 15. 参考パス一覧

```
C:\システム開発\工程管理AIプロジェクト_JAVA\
C:\システム開発\後加工工賃検証_JAVA\
\\192.168.0.101\共有フォルダ\国分工場\国分管理\後加工\後加工工賃明細\●自動検証\
  .cursor\skills\kouchin-consistency-check\
  .cursor\skills\kouchin-monthly-trend\
  docs\specs\2026-09-07-kouchin-monthly-trend-design.md
```
