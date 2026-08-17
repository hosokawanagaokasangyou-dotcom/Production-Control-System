# 操作ログと段階2後終了ゲート Implementation Plan

> **For implementers:** Execute this plan task-by-task. Use the `test-driven-development` skill for each task.

**Goal:** 配台重要操作を共有フォルダへ操作者別に90日残し、段階2後にローカル最新 Excel と加工計画が不一致なら7桁入力なしでは終了できないようにする。

**Architecture:** `OperatorActionLogStore` が共有 DATA/`操作ログ` へ NDJSON 追記。`Stage2IdentityCloseGate` が起動内フラグと再比較で要7桁を判定。`SevenDigitChallenge` が数字の生成・照合。専用タブと `MainShellController` 終了フックでつなぐ。

**Tech Stack:** Java 21、JavaFX、JUnit 5、Jackson、Maven（`code_java/mvnw.cmd`）

---

### Task 1: OperatorActionLogStore

**Files:**
- Create: `code_java/src/main/java/jp/co/pm/ai/desktop/config/OperatorActionLogStore.java`
- Modify: `code_java/src/main/java/jp/co/pm/ai/desktop/config/AppPaths.java`（`OPERATOR_ACTION_LOG_DIR_NAME` / `resolveOperatorActionLogRoot`）
- Test: `code_java/src/test/java/jp/co/pm/ai/desktop/config/OperatorActionLogStoreTest.java`

### Task 2: SevenDigitChallenge / Stage2IdentityCloseGate

**Files:**
- Create: `code_java/src/main/java/jp/co/pm/ai/desktop/ui/SevenDigitChallenge.java`
- Create: `code_java/src/main/java/jp/co/pm/ai/desktop/Stage2IdentityCloseGate.java`
- Test: `code_java/src/test/java/jp/co/pm/ai/desktop/ui/SevenDigitChallengeTest.java`
- Test: `code_java/src/test/java/jp/co/pm/ai/desktop/Stage2IdentityCloseGateTest.java`

### Task 3: ダイアログ・タブ・終了フック・記録

**Files:**
- Create: `SevenDigitChallengeDialog.java`, `OperatorActionLogTab.fxml`, `OperatorActionLogTabController.java`
- Modify: `MainShellTabId`, `MainShellTabLayoutDefaults`, `MainShell.fxml`, `MainShellController`, `ResultDispatchTableTabController`
