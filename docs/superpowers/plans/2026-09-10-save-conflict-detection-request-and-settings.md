# 保存前ファイル競合検知（依頼書・設定・プロファイル）Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 共通競合基盤を、依頼書入力（設定 JSON・受注転記・後加工マスタ）、ユーザープロファイル、グローバル設定パッケージ既定、Gemini 資格情報の明示保存に適用する。

**Architecture:** コア計画の `FingerprintBaseline` / `SaveConflictChecker` / `ConflictSaveDialog` / `SaveConflictGate` を再利用。秘匿ファイル（Gemini）は中身を要約に出さない。

**Tech Stack:** Java / JavaFX / JUnit 5 / Maven（`code_java/mvnw.cmd`）

**前提:** `2026-09-10-save-conflict-detection-core-attendance.md` 完了。配台系計画は並行可だが、共通クラスの API を変えないこと。

**仕様正本:** `docs/superpowers/specs/2026-09-10-save-conflict-detection-design.md`

---

## File map

| ファイル | 役割 |
|----------|------|
| Create: `.../io/conflict/RequestFormSettingsConflictDiffSummarizer.java` | 設定 JSON |
| Create: `.../io/conflict/JuchuWorkbookConflictDiffSummarizer.java` | 受注 Excel |
| Create: `.../io/conflict/PostProcessingMasterConflictDiffSummarizer.java` | 後加工アップロード xlsx |
| Create: `.../io/conflict/UserProfileConflictDiffSummarizer.java` | プロファイル JSON |
| Create: `.../io/conflict/PackageDefaultsConflictDiffSummarizer.java` | init_setting 群 |
| Create: `.../io/conflict/GeminiCredentialsConflictDiffSummarizer.java` | 秘匿要約 |
| Modify: `reconciliation/ReconciliationApp.java`（設定保存・転記入口） | ゲート |
| Modify: `ui/PostProcessingProductMasterEditorPane.java`（または実クラス名） | ゲート |
| Modify: `UserProfilesTabController.java` | ゲート |
| Modify: `GlobalSettingsTabController.java` | ゲート |
| Modify: `EnvTabController.java`（暗号化保存） | ゲート |

---

### Task 1: 依頼書入力設定 JSON

**Files:**
- Create: `RequestFormSettingsConflictDiffSummarizer.java` + Test（候補リスト件数・パス文字列の変化を業務語で）
- Modify: 設定の明示保存入口（`ReconciliationApp` 内の「JSONを直接編集」保存、工場出荷戻しの書き戻し）

- [ ] baseline: サマリ同フォルダの `request_form_input_settings.json`
- [ ] ComboBox 候補の**自動保存は対象外**（仕様）。明示保存ボタン／ダイアログ確定のみゲート
- [ ] Commit: `feat: 依頼書設定 JSON の保存前外部変更検知を追加`

---

### Task 2: 受注ファイルへの転記・更新

**Files:**
- Create: `JuchuWorkbookConflictDiffSummarizer.java` + Test  
  - 可能なら POI で「受注ﾌｧｲﾙ」シートの行数変化  
  - 無理なら「受注ブックの内容が変更されています」+ ファイル名 + ハッシュ先頭
- Modify: `ReconciliationApp` の「受注ファイルへ自動転記・更新」／一括転記の書込直前

- [ ] baseline: 転記対象の受注 Excel パス（設定または UI で選択中のファイル）。読込成功時に capture
- [ ] 原本フォルダ `PM_AI_REQUEST_FORM_ORIGINAL_DIR` 配下は読取専用ルール対象外の書込先ではないこと（転記先は受注ファイル側）を確認
- [ ] Commit: `feat: 受注転記の保存前外部変更検知を追加`

---

### Task 3: 後加工商品マスタ（アップロード用 xlsx）

**Files:**
- Create: `PostProcessingMasterConflictDiffSummarizer.java` + Test（行数・品番列の増減。POI または「ブックが変更」）
- Modify: 後加工マスタ編集ペインの「④ Excel保存」`saveUpload`

- [ ] baseline: ユーザー指定のアップロード用 xlsx
- [ ] 本番 `後加工商品マスタ.xlsx` を直接触らない既存方針は維持
- [ ] Commit: `feat: 後加工マスタアップロード保存の外部変更検知を追加`

---

### Task 4: ユーザープロファイル

**Files:**
- Create: `UserProfileConflictDiffSummarizer.java` + Test（プロファイル名キーの増減）
- Modify: `UserProfilesTabController.java` の `onSaveAction`

- [ ] baseline: 保存対象の `~/.pm-ai-desktop/user-profiles/*.json`（今回書くファイル）
- [ ] Commit: `feat: ユーザープロファイル保存の外部変更検知を追加`

---

### Task 5: グローバル設定パッケージ既定

**Files:**
- Create: `PackageDefaultsConflictDiffSummarizer.java` + Test（`session_defaults_*` / 列順 / 見出し別名のキー差分）
- Modify: `GlobalSettingsTabController.java` の `onSavePackageDefaultsAction`

- [ ] baseline: 書き出す各 `init_setting/*.json` をセット
- [ ] Commit: `feat: パッケージ既定保存の外部変更検知を追加`

---

### Task 6: Gemini 資格情報（秘匿）

**Files:**
- Create: `GeminiCredentialsConflictDiffSummarizer.java` + Test  
  - 要約は必ず「認証ファイルが更新または置換されています」のみ。平文・鍵・JSON 断片を含めない  
  - テストで出力に `api` / `key` / `token` らしき文字列が含まれないことも assert
- Modify: `EnvTabController.java` の `onEncryptGeminiCredentialsAction`（書込前）

- [ ] baseline: `GEMINI_CREDENTIALS_JSON` または `code/gemini_credentials.encrypted.json` の実際の書込先
- [ ] Commit: `feat: Gemini 資格情報保存の外部変更検知（秘匿要約）を追加`

---

### Task 7: 仕様完了条件の最終確認

- [ ] 仕様の対象画面が一覧どおり Guard 経由か grep で確認（`SaveConflictChecker.check` の呼び出し箇所）
- [ ] 自動保存パスに誤って入れていないか確認（`scheduleDesktopSessionSave` 付近に競合ゲートが無いこと）
- [ ] 常時 `FileChannel.lock` / 長寿命 `RandomAccessFile` ロックを新規追加していないこと

---

## Spec coverage（本計画）

依頼書設定・受注転記・後加工アップロード・プロファイル・パッケージ既定・Gemini。これとコア／配台計画を合わせて仕様の完了条件を満たす。
