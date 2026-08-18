# 同一化チェック履歴 実装計画

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 同一化チェック実行のたびに Excel＋加工計画 JSON のセットを操作者別フォルダへ最大20件保存し、新メインタブで閲覧できるようにする。

**状態:** Task 1〜7 実装完了（2026-08-18）

**Architecture:** `IdentityCheckHistoryStore` が共有 DATA 配下へセット保存・prune・一覧を担う。`AladdinEntryDispatchPlanIdentityCheck.evaluate` 成功時にスナップショットを書き出す。閲覧は `OperatorActionLogTab` と同型の操作者コンボ＋一覧タブ。

**Tech Stack:** Java 21 / JavaFX / Maven (`code_java`) / Jackson (`JsonTableIo`) / JUnit 5

**仕様正本:** `docs/specs/2026-08-18-identity-check-history-tab-design.md`

---

## ファイル構成

| ファイル | 責任 |
|----------|------|
| `AppPaths.java` | `IDENTITY_CHECK_HISTORY_DIR_NAME` / `resolveIdentityCheckHistoryRoot` |
| `IdentityCheckHistoryStore.java` | 保存・prune・一覧・meta 読込 |
| `IdentityCheckHistoryStoreTest.java` | Store 単体テスト |
| `AladdinEntryDispatchPlanIdentityCheck.java` | evaluate 成功時に Store 呼び出し |
| `AladdinEntryDispatchPlanIdentityCheckTest.java` | 保存フックのテスト（TempDir） |
| `MainShellTabId.java` | `IDENTITY_CHECK_HISTORY` |
| `MainShellTabLayoutDefaults.java` | flat 順・「その他」グループ |
| `IdentityCheckHistoryTab.fxml` | 閲覧 UI |
| `IdentityCheckHistoryTabController.java` | 一覧・Excel 開く・JSON 表表示 |
| `MainShell.fxml` / `MainShellController.java` | タブ配線 |

---

### Task 1: AppPaths に履歴ルートを追加

**Files:**
- Modify: `code_java/src/main/java/jp/co/pm/ai/desktop/config/AppPaths.java`
- Test: `code_java/src/test/java/jp/co/pm/ai/desktop/config/AppPathsIdentityCheckHistoryTest.java`（新規・最小）

- [ ] **Step 1: 失敗するテストを書く**

```java
@Test
void resolveIdentityCheckHistoryRoot_isSiblingOfSummaryWorkbook(@TempDir Path temp) {
    Map<String, String> ui = Map.of(
            AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
            temp.resolve("shared").resolve("サマリ.xlsx").toString());
    Path root = AppPaths.resolveIdentityCheckHistoryRoot(ui);
    assertEquals(temp.resolve("shared").resolve("同一化チェック履歴").normalize(), root.normalize());
}
```

- [ ] **Step 2: テスト実行（失敗を確認）**

Run: `cmd /c "mvnw.cmd -q test -Dtest=AppPathsIdentityCheckHistoryTest"`（`code_java` で）

Expected: コンパイル失敗またはメソッド未定義

- [ ] **Step 3: 実装**

`AppPaths` に操作ログと同様:

```java
public static final String IDENTITY_CHECK_HISTORY_DIR_NAME = "同一化チェック履歴";

public static Path resolveIdentityCheckHistoryRoot(Map<String, String> ui) {
    return siblingOfSummaryAiDispatchWorkbook(ui, IDENTITY_CHECK_HISTORY_DIR_NAME);
}
```

（`OPERATOR_ACTION_LOG_DIR_NAME` / `resolveOperatorActionLogRoot` の直後に置く）

- [ ] **Step 4: テスト通過を確認**

- [ ] **Step 5: Commit**

```bash
git add code_java/src/main/java/jp/co/pm/ai/desktop/config/AppPaths.java \
  code_java/src/test/java/jp/co/pm/ai/desktop/config/AppPathsIdentityCheckHistoryTest.java
git commit -m "feat: 同一化チェック履歴ルートの AppPaths を追加"
```

---

### Task 2: IdentityCheckHistoryStore（保存・prune・一覧）

**Files:**
- Create: `code_java/src/main/java/jp/co/pm/ai/desktop/io/IdentityCheckHistoryStore.java`
- Create: `code_java/src/test/java/jp/co/pm/ai/desktop/io/IdentityCheckHistoryStoreTest.java`

- [ ] **Step 1: 失敗するテストを書く**

```java
@Test
void saveSnapshot_writesSetAndPrunesTo20(@TempDir Path temp) throws Exception {
    Map<String, String> ui = Map.of(
            AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
            temp.resolve("shared").resolve("サマリ.xlsx").toString(),
            AppPaths.KEY_PM_AI_OPERATOR_USER, "テスト太郎");
    Path excel = temp.resolve("src.xlsx");
    Files.writeString(excel, "dummy"); // 実コピーはバイナリでよい。xlsx でなくても Store はバイトコピー
    // 本物の xlsx が必要なら最小 POI で作る。または Files.copy の対象として空ファイルで十分
    Files.write(excel, new byte[] {1, 2, 3});
    PlanInputTabularIo.TabularSheet tab =
            new PlanInputTabularIo.TabularSheet(
                    List.of("機械名", "依頼NO"), List.of(List.of("M1", "T1")));

    for (int i = 0; i < 22; i++) {
        Thread.sleep(5); // フォルダ名衝突回避（またはテスト用に時刻注入）
        IdentityCheckHistoryStore.save(
                ui, excel, tab, "ok", "配台計画と加工計画は同一", 0,
                excel, temp.resolve("plan.xlsx"));
    }
    Path opDir = IdentityCheckHistoryStore.resolveOperatorDir(ui, "テスト太郎");
    try (var s = Files.list(opDir)) {
        assertEquals(20, s.filter(Files::isDirectory).count());
    }
}
```

時刻衝突を避けるため、Store は `save(..., Instant now)` のテスト用オーバーロード、またはフォルダ名にミリ秒を含める（`yyyyMMdd-HHmmss-SSS`）。仕様の `yyyyMMdd-HHmmss` を守るなら、同一秒内は連番サフィックス `_2` を付ける。

推奨 API:

```java
public final class IdentityCheckHistoryStore {
    public static final int MAX_SNAPSHOTS_PER_USER = 20;
    public static final String EXCEL_FILE = "配台計画.xlsx";
    public static final String PLAN_JSON_FILE = "加工計画.json";
    public static final String META_FILE = "meta.json";

    public record Meta(
            String savedAt,
            String operator,
            String result,
            String badgeText,
            int diffCount,
            String excelSourcePath,
            String planSourcePath,
            String excelFileName,
            String planJsonFileName) {}

    public record SnapshotRef(Path dir, Meta meta) {}

    public static Path resolveRoot(Map<String, String> ui);
    public static Path resolveOperatorDir(Map<String, String> ui, String operator);
    /** @return 保存先ディレクトリ。失敗時 empty（例外は投げない） */
    public static Optional<Path> save(
            Map<String, String> ui,
            Path excelPath,
            PlanInputTabularIo.TabularSheet planTab,
            String result,
            String badgeText,
            int diffCount,
            Optional<Path> excelSourcePath,
            Optional<Path> planSourcePath);
    public static List<SnapshotRef> listNewestFirst(Map<String, String> ui, String operator);
    public static List<String> listOperatorDirNames(Map<String, String> ui);
    static void prune(Path operatorDir) throws IOException;
}
```

- [ ] **Step 2: テスト実行（失敗確認）**

- [ ] **Step 3: Store 実装**

実装要点:
- `Files.copy(excel, dest/配台計画.xlsx, REPLACE_EXISTING)`
- `JsonTableIo.saveArrayTable(dest/加工計画.json, headers, rows)`
- meta は Jackson `ObjectMapper`（`OperatorActionLogStore` と同様）で UTF-8 書き込み
- prune: 操作者直下のディレクトリを `lastModified` 昇順で削除し件数 ≤ 20
- `listNewestFirst`: meta.json があるディレクトリを新しい順
- パス安全: `OperatorActionLogStore.resolveOperatorDir` と同様に `startsWith(root)` チェック

- [ ] **Step 4: テスト通過**

追加ケース:
- `save` は excel 欠落時 empty
- 操作者ディレクトリが分離される

- [ ] **Step 5: Commit**

```bash
git commit -m "feat: 同一化チェック履歴の IdentityCheckHistoryStore を追加"
```

---

### Task 3: evaluate 成功時にスナップショット保存

**Files:**
- Modify: `code_java/src/main/java/jp/co/pm/ai/desktop/io/AladdinEntryDispatchPlanIdentityCheck.java`
- Modify: `code_java/src/test/java/jp/co/pm/ai/desktop/io/AladdinEntryDispatchPlanIdentityCheckTest.java`

- [ ] **Step 1: 失敗するテスト**

既存 `evaluate_identicalForOperatorGenerationAndMatchingPlan` の ui に  
`KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK` を TempDir 配下で設定し、evaluate 後に履歴フォルダに 1 セットあることを assert。

- [ ] **Step 2: evaluate 内で保存**

`compare` 成功後（`error == false` の Result を組み立てる直前）:

```java
String resultKey = compared.identical() ? "ok" : "mismatch";
IdentityCheckHistoryStore.save(
        u,
        excel.get(),
        tab,
        resultKey,
        compared.badgeText(),
        compared.diffs() != null ? compared.diffs().size() : 0,
        excel,
        planSource);
```

`tab` は完了行除外前の原本（設計: チェック時に読んだ表）。除外後 `activePlan` ではなく **読込直後の `tab`** を保存する。

保存失敗は握りつぶし（ログのみ。既存の `shell.appendLog` は Controller 側。Store 内で `System.err` や将来のログは使わず、戻り値 empty で無視）。

- [ ] **Step 3: テスト通過・Commit**

```bash
git commit -m "feat: 同一化チェック成功時に履歴セットを保存する"
```

---

### Task 4: MainShellTabId とレイアウト既定

**Files:**
- Modify: `code_java/src/main/java/jp/co/pm/ai/desktop/MainShellTabId.java`
- Modify: `code_java/src/main/java/jp/co/pm/ai/desktop/config/MainShellTabLayoutDefaults.java`

- [ ] **Step 1: enum 追加**

```java
IDENTITY_CHECK_HISTORY("identityCheckHistory"),
```

`OPERATOR_ACTION_LOG` の近くに置く。

- [ ] **Step 2: DEFAULT_FLAT_TAB_KEY_ORDER**

`OPERATOR_ACTION_LOG` の直後に `IDENTITY_CHECK_HISTORY.key()` を追加。

- [ ] **Step 3: groupedLayout「その他」**

`OPERATOR_ACTION_LOG` の直後に  
`MainShellTabLayoutNode.tabNode(MainShellTabId.IDENTITY_CHECK_HISTORY.key(), "")` を追加。

- [ ] **Step 4: Commit**

```bash
git commit -m "feat: 同一化チェック履歴タブ ID と既定レイアウトを追加"
```

---

### Task 5: 閲覧タブ FXML + Controller

**Files:**
- Create: `code_java/src/main/resources/jp/co/pm/ai/desktop/fxml/IdentityCheckHistoryTab.fxml`
- Create: `code_java/src/main/java/jp/co/pm/ai/desktop/IdentityCheckHistoryTabController.java`

- [ ] **Step 1: FXML**（`OperatorActionLogTab.fxml` を雛形）

- タイトル「同一化チェック履歴」
- 説明文（操作者別・最新20件・Excel＋加工計画JSON）
- 操作者 ComboBox / 再読み
- TableView: 日時 / 結果 / 差異件数 / フォルダ名
- ボタン: 「Excelを開く」「JSONを表で表示」

- [ ] **Step 2: Controller**

```java
public final class IdentityCheckHistoryTabController {
    public void bindShell(MainShellController shell);
    @FXML void onRefreshAction();
    @FXML void onOpenExcelAction();
    @FXML void onShowPlanJsonAction();
}
```

- 操作者一覧: `IdentityCheckHistoryStore.listOperatorDirNames` ＋ログイン中を先頭・選択
- 選択行の Excel: `Desktop.getDesktop().open(path)` または既存の Excel オープンヘルパーがあればそれを使う（`DispatchAladdinEntryGenerationDialog` / shell の open 系を grep）
- JSON: 簡易ダイアログで `JsonTableIo.loadArrayTable` → `TableView` 表示、または既存スプレッドシート支援があれば流用。最小は `Alert`/`Dialog` + `TableView`

- [ ] **Step 3: Commit**

```bash
git commit -m "feat: 同一化チェック履歴の閲覧タブ UI を追加"
```

---

### Task 6: MainShell 配線

**Files:**
- Modify: `code_java/src/main/resources/jp/co/pm/ai/desktop/fxml/MainShell.fxml`
- Modify: `code_java/src/main/java/jp/co/pm/ai/desktop/MainShellController.java`

- [ ] **Step 1: MainShell.fxml**

操作ログ Tab の直後に:

```xml
<Tab fx:id="mainShellTabIdentityCheckHistory" closable="false" text="同一化チェック履歴">
    <content>
        <fx:include fx:id="identityCheckHistoryTab" source="IdentityCheckHistoryTab.fxml"/>
    </content>
</Tab>
```

- [ ] **Step 2: MainShellController**

- `@FXML private Tab mainShellTabIdentityCheckHistory;`
- `@FXML private IdentityCheckHistoryTabController identityCheckHistoryTabController;`
- `bindShell` 呼び出し（他タブと同様の初期化箇所）
- `mainShellTabFor` / switch に `IDENTITY_CHECK_HISTORY` を追加（既存の `OPERATOR_ACTION_LOG` 分岐をコピー）

- [ ] **Step 3: コンパイル**

Run: `cmd /c "mvnw.cmd -q test-compile"`

- [ ] **Step 4: Commit**

```bash
git commit -m "feat: メインシェルに同一化チェック履歴タブを配線"
```

---

### Task 7: 手動確認チェックリスト

- [ ] アプリ起動 → 同一化チェック実行 → 共有 DATA の `同一化チェック履歴/{操作者}/` にセットが出来る
- [ ] 21 回以上で古いフォルダが消える（または単体テストで確認済みなら省略可）
- [ ] 閲覧タブで一覧・操作者切替・Excel オープン・JSON 表示
- [ ] 設計ドキュメントの状態を「実装完了」に更新（任意）

```bash
git commit -m "docs: 同一化チェック履歴タブ設計を実装完了に更新"
```

---

## Spec coverage

| 仕様 | Task |
|------|------|
| 毎回保存・セット内容 | 2, 3 |
| JSON はチェック時スナップショット | 3（`tab` 原本） |
| 共有 DATA / 操作者 / 20件 | 1, 2 |
| 閲覧タブ・操作者切替 | 5, 6 |
| エラー時はセット保存しない | 3（evaluate の成功パスのみ） |

## 実行手番

Plan complete and saved to `docs/specs/2026-08-18-identity-check-history-tab-plan.md`.

**実行方法を選んでください:**

1. **Subagent-Driven（推奨）** — タスクごとにサブエージェントを起動し、間でレビュー  
2. **Inline Execution** — このセッションで順に実装  

どちらで進めますか？
