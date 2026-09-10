# 保存前ファイル競合検知（配台・ルール・材料・ガント）Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 共通競合基盤を、配台マスタ・タスク入力・配台不要／特別ルール・材料表・設備ガントの明示保存に適用する。

**Architecture:** `docs/superpowers/plans/2026-09-10-save-conflict-detection-core-attendance.md` の `FingerprintBaseline` / `SaveConflictChecker` / `ConflictSaveDialog` / `SaveConflictGate` を再利用。画面ごとに Summarizer と baseline パスを定義し、既存の保存入口の直前でゲートする。

**Tech Stack:** Java / JavaFX / JUnit 5 / Maven（`code_java/mvnw.cmd`）

**前提:** コア＋勤怠系計画がマージ済みであること。

**仕様正本:** `docs/superpowers/specs/2026-09-10-save-conflict-detection-design.md`

---

## File map

| ファイル | 役割 |
|----------|------|
| Create: `.../io/conflict/JsonStructureConflictDiffSummarizer.java` | JSON キー差分の汎用＋画面ラベル付き |
| Create: `.../io/conflict/MasterDispatchConflictDiffSummarizer.java` | 配台マスタ業務語 |
| Create: `.../io/conflict/PlanInputConflictDiffSummarizer.java` | タスク入力表 |
| Create: `.../io/conflict/ExcludeRulesConflictDiffSummarizer.java` | 配台不要ルール |
| Create: `.../io/conflict/SpecialRulesConflictDiffSummarizer.java` | 特別ルール JSON |
| Create: `.../io/conflict/LookupTablesConflictDiffSummarizer.java` | 材料 .txt |
| Create: `.../io/conflict/EquipmentGanttConflictDiffSummarizer.java` | ガント担当 |
| Modify: `MasterDispatchSheetsTabController.java` | ゲート |
| Modify: `PlanInputTabController.java` | ゲート |
| Modify: `ExcludeRulesTabController.java` | ゲート |
| Modify: `SpecialRulesTabController.java` | ゲート |
| Modify: `dispatch/rules/SpecialRulesBuilderTabController.java` | ゲート |
| Modify: `dispatch/rules/ProcessMachinePriorityTabController.java` | ゲート |
| Modify: `CodeDispatchLookupTablesTabController.java`（および子パネルの save） | ゲート |
| Modify: `EquipmentGanttGraphicTabController.java` | ゲート |

各 Summarizer に対応する `*Test.java` を `code_java/src/test/java/jp/co/pm/ai/desktop/io/conflict/` に置く。

---

### Task 1: JsonStructureConflictDiffSummarizer（ルール系の土台）

**Files:**
- Create: `code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/JsonStructureConflictDiffSummarizer.java`
- Create: `code_java/src/test/java/jp/co/pm/ai/desktop/io/conflict/JsonStructureConflictDiffSummarizerTest.java`

- [ ] **Step 1: テスト** — トップレベルキーの追加削除、配列長の変化を「・キー X が追加」「・配列 Y の要素数が 2→5」と日本語で出す。壊れた JSON はフォールバック。

```java
@Test
void reportsAddedKeyAndArraySizeChange() throws Exception {
    Path p = Path.of("rules.json").toAbsolutePath().normalize();
    byte[] base = "{\"rules\":[1],\"meta\":{}}".getBytes(StandardCharsets.UTF_8);
    byte[] disk = "{\"rules\":[1,2,3],\"meta\":{},\"newKey\":true}".getBytes(StandardCharsets.UTF_8);
    String s =
            new JsonStructureConflictDiffSummarizer("配台不要ルール")
                    .summarize(Map.of(p, base), Map.of(p, disk), List.of(p));
    assertTrue(s.contains("newKey"), s);
    assertTrue(s.contains("rules"), s);
}
```

- [ ] **Step 2: 実装** — Jackson でツリー比較。深さは 2 までに制限（長文防止）。プレフィックスに画面名。

- [ ] **Step 3: テスト PASS → Commit** `feat: JSON 構造差分の競合要約ユーティリティを追加`

---

### Task 2: 配台マスタ

**Files:**
- Create: `MasterDispatchConflictDiffSummarizer.java` + Test
- Modify: `MasterDispatchSheetsTabController.java`

- [ ] baseline: `AppPaths` の master-dispatch-sheets JSON + `PM_AI_MASTER_WORKBOOK`（master.xlsm）
- [ ] Summarizer: シート名（skills / need / speed 等）ごとの行数変化・「格子が変更」件数。xlsm は「マスタブックが変更されています」
- [ ] `saveCurrentGridsToDisk`（または `onSaveAction`）の書込前にゲート。タイトル「配台マスタ」
- [ ] 読込成功・保存成功で `FingerprintBaseline.capture`
- [ ] テスト + コンパイル → Commit: `feat: 配台マスタの保存前外部変更検知を追加`

---

### Task 3: 配台計画（タスク入力）

**Files:**
- Create: `PlanInputConflictDiffSummarizer.java` + Test
- Modify: `PlanInputTabController.java`

- [ ] baseline: 現在の `PM_AI_PLAN_INPUT_PATH` が指す単一ファイル（csv/xlsx/xlsm）
- [ ] Summarizer:  
  - csv: 行数変化・先頭数列の差分件数  
  - xlsx/xlsm: 「ブック内容が変更されています」（必要なら Apache POI でシート行数のみ）
- [ ] `onSaveButtonAction` の実保存前にゲート。タイトル「配台計画（タスク入力）」
- [ ] Commit: `feat: タスク入力の保存前外部変更検知を追加`

---

### Task 4: 配台不要ルール

**Files:**
- Create: `ExcludeRulesConflictDiffSummarizer.java`（内部で `JsonStructureConflictDiffSummarizer` を委譲し、ルール件数を「配台不要ルールが N→M 件」と添える）+ Test
- Modify: `ExcludeRulesTabController.java`

- [ ] baseline: `PM_AI_EXCLUDE_RULES_JSON` 解決パスのみ
- [ ] `onSaveButtonAction` 前にゲート
- [ ] Commit: `feat: 配台不要ルールの保存前外部変更検知を追加`

---

### Task 5: 特別ルール（JSON／ビルダー／工程優先）

**Files:**
- Create: `SpecialRulesConflictDiffSummarizer.java` + Test（ルール件数・`processMachinePriorities` の工程数）
- Modify: `SpecialRulesTabController.java`（`onSaveJsonAction`）
- Modify: `dispatch/rules/SpecialRulesBuilderTabController.java`（`onSaveAction`）
- Modify: `dispatch/rules/ProcessMachinePriorityTabController.java`（`saveToDisk`）

- [ ] 3 入口とも同じ `dispatch_special_rules.json` パスの baseline（シェルまたは `DispatchRulePaths` の既存解決を使う）
- [ ] いずれかで保存成功したら、可能なら他コントローラの baseline も更新する（同一シェル参照があれば）。無ければ次回保存で競合→再読込（仕様どおり）
- [ ] Commit: `feat: 特別ルール系の保存前外部変更検知を追加`

---

### Task 6: 材料・製品種類（.txt 群）

**Files:**
- Create: `LookupTablesConflictDiffSummarizer.java` + Test（キー行の追加削除を「・キー〇〇が追加」）
- Modify: `CodeDispatchLookupTablesTabController.java` および実際に `saveToDisk` する子クラス

- [ ] baseline: 保存対象の各 `.txt` をセット指紋
- [ ] 各 `saveToDisk` 入口で、そのファイル単体またはセット全体をチェック（実装は「今回書くファイル＋同一フォルダの既知セット」）。最低限「今回上書きする Path」は必須
- [ ] Commit: `feat: 材料・製品種類表の保存前外部変更検知を追加`

---

### Task 7: 設備ガント担当保存

**Files:**
- Create: `EquipmentGanttConflictDiffSummarizer.java` + Test（担当者の追加削除をバー ID／機械名で）
- Modify: `EquipmentGanttGraphicTabController.java`（`beginSaveAssignmentChanges`）

- [ ] baseline: 読込中の計画 JSON + 契約 JSON（`…設.json` / `*_equipment_gantt_contract.json`）。xlsx 同期する場合はセットに含める
- [ ] Commit: `feat: 設備ガント担当保存の外部変更検知を追加`

---

### Task 8: 手動確認

- [ ] 各画面で対象ファイルを外部編集 → 保存 → ダイアログ → 再読込／上書き／キャンセル

---

## Spec coverage（本計画）

配台マスタ・タスク入力・exclude／special／lookup／gantt の明示保存。勤怠・依頼書・プロファイルは他計画。
