# TPI PDF 一覧フィルタ Implementation Plan

> **For implementers:** Execute this plan task-by-task. Use TDD for filter logic.

**Goal:** 依頼書入力の一覧ラジオに「TPI PDF」を追加し、TPI 由来／ステータスに「TPI PDF」を含む行だけ表示する。

**Architecture:** 既存 `RecordListFilterMode` + `recordIncludedInListFilter` にモード追加。UI は `ReconciliationApp` のラジオ群。

**Tech Stack:** Java / JavaFX

---

### Task 1: フィルタ判定（TDD）

**Files:**
- Modify: `RecordListFilterMode.java`
- Modify: `ReconciliationApp.java`（`isTpiRelatedRecord` / `recordIncludedInListFilter`）
- Modify: `ReconciliationAppRecordFilterTest.java`

- [ ] 失敗テストを書く（TPI_PDF raw / ステータス含有 / 非 TPI 除外）
- [ ] `TPI_PDF` enum と判定ロジックを実装して通す

### Task 2: UI ラジオ

**Files:**
- Modify: `ReconciliationApp.java`

- [ ] `rbTpiPdfFilter` 追加、`resolveRecordListFilterMode` 接続
- [ ] TPI 無効時は非表示（`AppPaths.isRequestFormTpiPdfEnabled`）
- [ ] レイアウトを 3 行目に配置

### Task 3: 確認

- [ ] `mvnw test -Dtest=ReconciliationAppRecordFilterTest`
