# remote_log Implementation Plan

> **For agentic workers:** Implement task-by-task. Steps use checkbox syntax.

**Goal:** 段階1／2／2.1 終了時に共有 `remote_log/<操作者>/` へ UI・Python ログを3日世代管理で保存する。

**Architecture:** `RemoteSupportLogArchive` がパス解決・書込・削除を担当。`MainShellController.completeStageRunOnFx` 末尾から非同期呼び出し。

**Tech Stack:** Java 26 / JavaFX / JUnit 5

---

## Task 1: AppPaths + EnvVarDocs

- [ ] `KEY_PM_AI_REMOTE_LOG`, `REMOTE_LOG_DIR_NAME`, `resolveRemoteLogRoot`, `resolveExecutionLogTxtPath`
- [ ] `EnvVarDocs` に説明追加

## Task 2: RemoteSupportLogArchive + テスト

- [ ] 純ロジック（世代名、保持削除、無効判定）の単体テスト
- [ ] 実装

## Task 3: UI 接続

- [ ] `MainRunTabController` でログ全文スナップショット
- [ ] `completeStageRunOnFx` から stage1/2/2.1 で呼び出し
