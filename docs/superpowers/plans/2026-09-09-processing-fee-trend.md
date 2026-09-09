# 加工賃トレンド Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** メインシェルに独立タブ「加工賃トレンド」を追加し、受注 AH × 加工日 m の実績円・予定円・累計折れ線を表示する。

**Architecture:** 円専用の読込（AH マップ）＋集計（`ProcessingFeeTrendAggregator`）＋タブ UI。既存 m の `ProcessingTrendAggregator.DayPoint` は変更しない。日次 m×単価は依頼No 結合。チャートは棒2本＋累計 Path（m トレンドの Path オーバーレイ方針に倣う）。

**Tech Stack:** Java 21 / JavaFX / JUnit 5 / Maven（`code_java/mvnw.cmd`）／Apache POI（受注読込）

**仕様正本:** [docs/specs/2026-09-09-processing-fee-trend-design.md](../specs/2026-09-09-processing-fee-trend-design.md)

---

## File map

| ファイル | 役割 |
|----------|------|
| Create: `.../io/actuals/JuchuProcessingFeeRateLoader.java` | 受注ﾌｧｲﾙから依頼No→AH（円/m）Map |
| Create: `.../io/actuals/ProcessingFeeTrendAggregator.java` | 日次実績円・予定円・累計 |
| Create: `.../ProcessingFeeTrendTabController.java` + FXML + CSS | 独立タブ UI |
| Create: tests for loader / aggregator / FXML |
| Modify: `MainShellTabId`, `MainShellTabLayoutDefaults`, `MainShell.fxml`, `MainShellController` | タブ登録 |
| Modify: `EquipmentStatusDashboardSourceLoader` または専用解決 | 受注パス／キャッシュ読取（原本は読取専用） |

---

### Task 1: AH 単価ローダ（TDD）

**Files:**
- Create: `code_java/src/test/java/jp/co/pm/ai/desktop/io/actuals/JuchuProcessingFeeRateLoaderTest.java`
- Create: `code_java/src/main/java/jp/co/pm/ai/desktop/io/actuals/JuchuProcessingFeeRateLoader.java`

- [ ] **Step 1: 失敗するテスト** — 一時 xlsx（シート名 `受注ﾌｧｲﾙ`、見出し行3、A=依頼No、AH=加工賃）から `Map<String,Double>` が取れること。空 AH はマップに載せないか 0 を仕様どおり。
- [ ] **Step 2: テスト実行 → 失敗確認**
- [ ] **Step 3: 実装** — 列 index は `JuchuSheetColumnLayout.Col.IRAI_NO` / `KAKOCHIN`。見出し行はレジストリ既定（3）を引数または定数で。数値パース（カンマ除去）。複数行同一依頼No は **後勝ちまたは最大** をテストで固定。
- [ ] **Step 4: テスト成功**
- [ ] **Step 5: commit** — `feat: 受注AH加工賃単価ローダを追加`

---

### Task 2: 円集計 Aggregator（TDD）

**Files:**
- Create: `.../ProcessingFeeTrendAggregatorTest.java`
- Create: `.../ProcessingFeeTrendAggregator.java`

- [ ] **Step 1: 失敗するテスト** — 入力として「日付・依頼No・m」の行リスト＋単価 Map を与え、日次 `actualYen` / `planYen`、累計 `actualCumYen` / `planCumYen` が正しいこと。
- [ ] **Step 2: 失敗確認**
- [ ] **Step 3: 実装** — `DayPoint(date, actualYen, planYen, actualCumYen, planCumYen)` と `Result`。期間の全日 0 埋め。実績累計は当日まで（m 側と同様）。予定累計は期間全体。
- [ ] **Step 4: AH 欠落行は 0 円になるテスト**
- [ ] **Step 5: commit** — `feat: 加工賃トレンド日次・累計集計を追加`

**Note:** 既存 `ProcessingTrendAggregator` の内部行（依頼NO・数量）を再利用できるなら、private 行抽出を package API 化するより、Fee 側でソース行を再走査する方が結合が明確。どちらでもテストで契約を固定。

---

### Task 3: メインシェルタブ骨組み

**Files:**
- Modify: `MainShellTabId.java`, `MainShellTabLayoutDefaults.java`, `MainShell.fxml`, `MainShellController.java`
- Create: `ProcessingFeeTrendTab.fxml`, `ProcessingFeeTrendTabController.java`（空状態＋プレースホルダ）
- Create: `ProcessingFeeTrendTabFxmlTest.java`
- CSS: `pm-ai-desktop.css` に `pm-processing-fee-trend-*`（必要最小）

- [ ] **Step 1: タブ ID・DEFAULT 末尾追加・groupedLayout・FXML Tab・Controller 配線**
- [ ] **Step 2: FXML ロードテスト**
- [ ] **Step 3: commit** — `feat: 加工賃トレンドタブをメインシェルに追加`

---

### Task 4: チャート・KPI・読込配線

**Files:**
- Modify: `ProcessingFeeTrendTabController.java`
- 受注パス解決: `AppPaths` / 既存依頼書・受注解決を調査して再利用
- 実績・予定: `EquipmentStatusDashboardSourceLoader.LoadedSources` を加工トレンドと同様に取得

- [ ] **Step 1: BG 読込 → Fee Aggregator → KPI（実績円合計・予定円合計）**
- [ ] **Step 2: BarChart（実績・予定）＋ Path 累計（実績累計・予定累計）。既存 m の Path オーバーレイを参考**
- [ ] **Step 3: 期間プリセット最小セット（今月／先月／直近3ヶ月）**
- [ ] **Step 4: 手動確認チェックリストを PR／コミットメッセージに記載**
- [ ] **Step 5: commit** — `feat: 加工賃トレンドに実績・予定・累計グラフを接続`

---

### Task 5: Excel 出力（任意・初版に含める場合）

- [ ] `.pm-ai-cache/exports/processing-fee-trend/` へ日次シート出力
- [ ] commit — `feat: 加工賃トレンド Excel 出力`

（時間がなければ Task 4 までで一度出荷し、Excel は後続。）

---

## 手動確認

1. タブ「加工賃トレンド」が表示される
2. データあり期間で棒（実績・予定）と累計折れ線が出る
3. m トレンドタブの表示が変わっていない
4. AH 空の依頼は金額 0（またはサマリ警告）
5. 受注原本フォルダに書き込みが無い

## 完了条件

- 上記テスト緑
- design spec の決定事項を満たす
- `MainShellTabLayoutDefaults` に新キー明示
