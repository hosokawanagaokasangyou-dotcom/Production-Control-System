# ローカル最新・世代を開くボタン Implementation Plan

> **For implementers:** Execute this plan task-by-task. Use the `test-driven-development` skill for each task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 結果_配台表タブに「ローカル最新を開く」「ローカル世代を開く…」を追加し、ローカル出力先の xlsx／世代を開けるようにする。

**Architecture:** 共有側ハンドラを鏡像する。`DispatchAladdinEntryGenerationDialog` にルート解決（SHARED／LOCAL）を渡し、既存 `show` は SHARED 互換ラッパにする。FXML は「ローカルへ出力」直後に 2 ボタンを置く。

**Tech Stack:** JavaFX FXML、JUnit 5、Maven（`code_java`）

**仕様正本:** `docs/specs/2026-07-14-aladdin-local-open-buttons-design.md`

---

### Task 1: 世代ルート解決ヘルパ

**Files:**
- Modify: `code_java/src/main/java/jp/co/pm/ai/desktop/ui/DispatchAladdinEntryGenerationDialog.java`
- Create: `code_java/src/test/java/jp/co/pm/ai/desktop/ui/DispatchAladdinEntryGenerationDialogTest.java`

- [ ] 失敗テスト: `generationRoot(ui, SHARED)` / `LOCAL` が `AppPaths` の共有／ローカル dir と一致
- [ ] `static Path generationRoot(Map, Destination)` を実装し、`show` がそれを使う
- [ ] コミット

### Task 2: ダイアログ show オーバーロード + UI ボタン

**Files:**
- Modify: `DispatchAladdinEntryGenerationDialog.java`（タイトル・rootLabel）
- Modify: `ResultDispatchTableTab.fxml`
- Modify: `ResultDispatchTableTabController.java`

- [ ] `show(owner, ui, defaultOperator, Destination)` 追加。既存 3 引数は SHARED 委譲
- [ ] FXML に「ローカル最新を開く」「ローカル世代を開く…」
- [ ] Controller ハンドラ（ローカル最新／ローカル世代）
- [ ] コミット

### Task 3: 仕様ステータス更新

- [ ] 設計メモを「実装完了」に
- [ ] コミット・push
