# 受注検索タブ Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 依頼書入力に「受注検索」子タブを追加し、読込済み受注行を納期範囲（希望／調整 OR）と製品名／投入原反（部分一致 OR）で閲覧検索できるようにする。

**Architecture:** 検索条件・判定は UI 非依存の `JuchuOrderSearchCriteria` / `JuchuOrderSearch`。UI は `JuchuOrderSearchPane`。`ReconciliationApp` はタブ登録と `orderRecords` の供給のみ。`MainShellInnerTabCatalog` の並びと、マスター一覧の nested index を更新する。

**Tech Stack:** Java / JavaFX / JUnit 5 / Maven（`code_java/mvnw.cmd`）

**仕様正本:** `docs/superpowers/specs/2026-09-07-juchu-order-search-tab-design.md`

---

## File map

| ファイル | 役割 |
|----------|------|
| Create: `code_java/src/main/java/jp/co/pm/ai/desktop/reconciliation/JuchuOrderSearchCriteria.java` | From/To・キーワード + バリデーション |
| Create: `code_java/src/main/java/jp/co/pm/ai/desktop/reconciliation/JuchuOrderSearch.java` | `matches` / `filter` 純関数 |
| Create: `code_java/src/test/java/jp/co/pm/ai/desktop/reconciliation/JuchuOrderSearchTest.java` | 検索ロジック単体テスト |
| Create: `code_java/src/main/java/jp/co/pm/ai/desktop/reconciliation/JuchuOrderSearchPane.java` | 左条件・右結果 UI |
| Create: `code_java/src/test/java/jp/co/pm/ai/desktop/config/MainShellInnerTabCatalogRequestFormInputTest.java` | 子タブラベル順・nested index |
| Modify: `code_java/src/main/java/jp/co/pm/ai/desktop/config/MainShellInnerTabCatalog.java` | labels + nested index 5 |
| Modify: `code_java/src/main/java/jp/co/pm/ai/desktop/reconciliation/ReconciliationApp.java` | タブ追加・データ供給 |

---

### Task 1: 検索ロジック（TDD）

**Files:**
- Create: `code_java/src/test/java/jp/co/pm/ai/desktop/reconciliation/JuchuOrderSearchTest.java`
- Create: `code_java/src/main/java/jp/co/pm/ai/desktop/reconciliation/JuchuOrderSearchCriteria.java`
- Create: `code_java/src/main/java/jp/co/pm/ai/desktop/reconciliation/JuchuOrderSearch.java`

- [ ] **Step 1: 失敗するテストを書く**

```java
package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import org.junit.jupiter.api.Test;

class JuchuOrderSearchTest {

    private static OrderRecord rec(String reqNo, Map<String, String> db) {
        return new OrderRecord(reqNo, "既存", "U", "", "", Map.of(), db);
    }

    @Test
    void validation_requiresDateRangeAndAtLeastOneKeyword() {
        assertTrue(
                new JuchuOrderSearchCriteria(null, LocalDate.of(2026, 6, 1), "A", "")
                        .validationError()
                        .isPresent());
        assertTrue(
                new JuchuOrderSearchCriteria(LocalDate.of(2026, 6, 1), null, "A", "")
                        .validationError()
                        .isPresent());
        assertTrue(
                new JuchuOrderSearchCriteria(
                                LocalDate.of(2026, 6, 2), LocalDate.of(2026, 6, 1), "A", "")
                        .validationError()
                        .isPresent());
        assertTrue(
                new JuchuOrderSearchCriteria(
                                LocalDate.of(2026, 6, 1), LocalDate.of(2026, 6, 30), "", "  ")
                        .validationError()
                        .isPresent());
        assertEquals(
                Optional.empty(),
                new JuchuOrderSearchCriteria(
                                LocalDate.of(2026, 6, 1), LocalDate.of(2026, 6, 30), "製品X", "")
                        .validationError());
    }

    @Test
    void matches_hopeDeliveryInRange() {
        var c =
                new JuchuOrderSearchCriteria(
                        LocalDate.of(2026, 6, 1), LocalDate.of(2026, 6, 30), "ABC", "");
        assertTrue(
                JuchuOrderSearch.matches(
                        rec("1", Map.of("希望納期", "2026-06-15", "製品", "XXABCXX")), c));
        assertFalse(
                JuchuOrderSearch.matches(
                        rec("2", Map.of("希望納期", "2026-07-01", "製品", "XXABCXX")), c));
    }

    @Test
    void matches_adjustDeliveryInRange_evenIfHopeOutside() {
        var c =
                new JuchuOrderSearchCriteria(
                        LocalDate.of(2026, 6, 1), LocalDate.of(2026, 6, 30), "ABC", "");
        assertTrue(
                JuchuOrderSearch.matches(
                        rec(
                                "1",
                                Map.of(
                                        "希望納期",
                                        "2026-05-01",
                                        "調整納期",
                                        "2026-06-10",
                                        "製品",
                                        "ABC")),
                        c));
    }

    @Test
    void matches_productOrRawMaterial_or() {
        var c =
                new JuchuOrderSearchCriteria(
                        LocalDate.of(2026, 6, 1), LocalDate.of(2026, 6, 30), "PROD", "RAW");
        assertTrue(
                JuchuOrderSearch.matches(
                        rec("1", Map.of("希望納期", "2026-06-10", "製品", "xxPRODyy", "品名1", "zzz")),
                        c));
        assertTrue(
                JuchuOrderSearch.matches(
                        rec("2", Map.of("希望納期", "2026-06-10", "製品", "zzz", "原反品名", "xxRAWyy")),
                        c));
        assertFalse(
                JuchuOrderSearch.matches(
                        rec("3", Map.of("希望納期", "2026-06-10", "製品", "zzz", "品名1", "zzz")),
                        c));
    }

    @Test
    void matches_unparseableDelivery_doesNotHitOnDate() {
        var c =
                new JuchuOrderSearchCriteria(
                        LocalDate.of(2026, 6, 1), LocalDate.of(2026, 6, 30), "ABC", "");
        assertFalse(
                JuchuOrderSearch.matches(
                        rec("1", Map.of("希望納期", "不明", "調整納期", "", "製品", "ABC")), c));
    }

    @Test
    void filter_returnsMatchingOnly() {
        var c =
                new JuchuOrderSearchCriteria(
                        LocalDate.of(2026, 6, 1), LocalDate.of(2026, 6, 30), "HIT", "");
        List<OrderRecord> out =
                JuchuOrderSearch.filter(
                        List.of(
                                rec("a", Map.of("希望納期", "2026-06-05", "製品", "HIT")),
                                rec("b", Map.of("希望納期", "2026-06-05", "製品", "MISS"))),
                        c);
        assertEquals(1, out.size());
        assertEquals("a", out.get(0).getReqNo());
    }
}
```

- [ ] **Step 2: テスト実行（失敗確認）**

```bat
cd code_java
.\mvnw.cmd -q test -Dtest=JuchuOrderSearchTest
```

Expected: コンパイルエラー（クラス未定義）または FAIL

- [ ] **Step 3: 最小実装**

`JuchuOrderSearchCriteria.java`:

```java
package jp.co.pm.ai.desktop.reconciliation;

import java.time.LocalDate;
import java.util.Optional;

public record JuchuOrderSearchCriteria(
        LocalDate from, LocalDate to, String productKeyword, String rawMaterialKeyword) {

    public Optional<String> validationError() {
        if (from == null || to == null) {
            return Optional.of("納期範囲を指定してください");
        }
        if (from.isAfter(to)) {
            return Optional.of("開始日が終了日より後です");
        }
        if (isBlank(productKeyword) && isBlank(rawMaterialKeyword)) {
            return Optional.of("製品名または投入原反を入力してください");
        }
        return Optional.empty();
    }

    private static boolean isBlank(String s) {
        return s == null || s.isBlank();
    }
}
```

`JuchuOrderSearch.java`:

```java
package jp.co.pm.ai.desktop.reconciliation;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Map;
import java.util.Objects;

public final class JuchuOrderSearch {

    private JuchuOrderSearch() {}

    public static List<OrderRecord> filter(
            Collection<OrderRecord> records, JuchuOrderSearchCriteria criteria) {
        Objects.requireNonNull(criteria, "criteria");
        if (criteria.validationError().isPresent()) {
            throw new IllegalArgumentException(criteria.validationError().get());
        }
        List<OrderRecord> out = new ArrayList<>();
        if (records == null) {
            return out;
        }
        for (OrderRecord r : records) {
            if (matches(r, criteria)) {
                out.add(r);
            }
        }
        return out;
    }

    public static boolean matches(OrderRecord record, JuchuOrderSearchCriteria criteria) {
        if (record == null || criteria == null) {
            return false;
        }
        if (!deliveryInRange(record.getDbValues(), criteria.from(), criteria.to())) {
            return false;
        }
        return keywordMatches(record.getDbValues(), criteria);
    }

    private static boolean deliveryInRange(Map<String, String> db, LocalDate from, LocalDate to) {
        if (db == null) {
            return false;
        }
        return dateInRange(db.get("希望納期"), from, to) || dateInRange(db.get("調整納期"), from, to);
    }

    private static boolean dateInRange(String raw, LocalDate from, LocalDate to) {
        LocalDate d = JuchuTransferValueNormalizer.parseLocalDate(raw);
        if (d == null) {
            return false;
        }
        return !d.isBefore(from) && !d.isAfter(to);
    }

    private static boolean keywordMatches(Map<String, String> db, JuchuOrderSearchCriteria c) {
        boolean productOk =
                !isBlank(c.productKeyword()) && textContains(dbGet(db, "製品"), c.productKeyword());
        boolean rawOk =
                !isBlank(c.rawMaterialKeyword())
                        && (textContains(dbGet(db, "品名1"), c.rawMaterialKeyword())
                                || textContains(dbGet(db, "原反品名"), c.rawMaterialKeyword()));
        return productOk || rawOk;
    }

    private static String dbGet(Map<String, String> db, String key) {
        if (db == null) {
            return "";
        }
        String v = db.get(key);
        return v != null ? v : "";
    }

    private static boolean textContains(String haystack, String needle) {
        String h = JuchuTransferValueNormalizer.normalizeText(haystack);
        String n = JuchuTransferValueNormalizer.normalizeText(needle);
        return !n.isEmpty() && h.contains(n);
    }

    private static boolean isBlank(String s) {
        return s == null || s.isBlank();
    }

    /** 結果表の原反表示用（品名1 優先、空なら 原反品名）。 */
    public static String displayRawMaterial(Map<String, String> db) {
        String a = dbGet(db, "品名1").strip();
        if (!a.isEmpty()) {
            return a;
        }
        return dbGet(db, "原反品名").strip();
    }
}
```

- [ ] **Step 4: テスト再実行（成功確認）**

```bat
cd code_java
.\mvnw.cmd -q test -Dtest=JuchuOrderSearchTest
```

Expected: BUILD SUCCESS / tests pass

- [ ] **Step 5: Commit**

```bash
git add code_java/src/main/java/jp/co/pm/ai/desktop/reconciliation/JuchuOrderSearchCriteria.java \
  code_java/src/main/java/jp/co/pm/ai/desktop/reconciliation/JuchuOrderSearch.java \
  code_java/src/test/java/jp/co/pm/ai/desktop/reconciliation/JuchuOrderSearchTest.java
git commit -m "feat: 受注検索の納期・製品・原反フィルタロジックを追加"
```

---

### Task 2: MainShellInnerTabCatalog

**Files:**
- Modify: `code_java/src/main/java/jp/co/pm/ai/desktop/config/MainShellInnerTabCatalog.java`
- Create: `code_java/src/test/java/jp/co/pm/ai/desktop/config/MainShellInnerTabCatalogRequestFormInputTest.java`

- [ ] **Step 1: 失敗するテストを書く**

```java
package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.List;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.MainShellTabId;

class MainShellInnerTabCatalogRequestFormInputTest {

    @Test
    void labelsFor_requestFormInput_includesJuchuSearchAfterIndexSheet() {
        assertEquals(
                List.of(
                        "一括照合データベース・受注管理",
                        "目次シート",
                        "受注検索",
                        "【設定】",
                        "後加工商品マスタ",
                        "マスター一覧"),
                MainShellInnerTabCatalog.labelsFor(MainShellTabId.REQUEST_FORM_INPUT));
    }

    @Test
    void nestedInnerTabs_masterList_isIndex5() {
        assertEquals(
                List.of("機械コード", "工程マスタ", "加工内容マスタ"),
                MainShellInnerTabCatalog.nestedInnerTabLabelsUnderInnerTab(
                        MainShellTabId.REQUEST_FORM_INPUT, 5));
        assertEquals(
                List.of(),
                MainShellInnerTabCatalog.nestedInnerTabLabelsUnderInnerTab(
                        MainShellTabId.REQUEST_FORM_INPUT, 4));
    }
}
```

- [ ] **Step 2: テスト実行（失敗確認）**

```bat
cd code_java
.\mvnw.cmd -q test -Dtest=MainShellInnerTabCatalogRequestFormInputTest
```

Expected: FAIL（「受注検索」欠落 / nested index 不一致）

- [ ] **Step 3: カタログ更新**

`labelsFor` の `REQUEST_FORM_INPUT` を次に変更:

```java
case REQUEST_FORM_INPUT ->
        List.of(
                "一括照合データベース・受注管理",
                "目次シート",
                "受注検索",
                "【設定】",
                "後加工商品マスタ",
                "マスター一覧");
```

`nestedInnerTabLabelsUnderInnerTab` の条件を `innerTabIndex == 4` から `innerTabIndex == 5` に変更（マスター一覧が 5 番になるため）。

- [ ] **Step 4: テスト成功確認**

```bat
cd code_java
.\mvnw.cmd -q test -Dtest=MainShellInnerTabCatalogRequestFormInputTest
```

- [ ] **Step 5: Commit**

```bash
git add code_java/src/main/java/jp/co/pm/ai/desktop/config/MainShellInnerTabCatalog.java \
  code_java/src/test/java/jp/co/pm/ai/desktop/config/MainShellInnerTabCatalogRequestFormInputTest.java
git commit -m "feat: 依頼書入力の子タブに受注検索をカタログ登録"
```

---

### Task 3: JuchuOrderSearchPane（UI）

**Files:**
- Create: `code_java/src/main/java/jp/co/pm/ai/desktop/reconciliation/JuchuOrderSearchPane.java`

- [ ] **Step 1: Pane 実装**

左条件・右結果の `BorderPane` / `SplitPane` を返す。既存 CSS クラス（`form-scroll-pane` 等）を流用し、スタイルは最小限。

要点:

```java
public final class JuchuOrderSearchPane {
    public static Parent build(Supplier<List<OrderRecord>> recordsSupplier) {
        // 左: DatePicker from/to, TextField product, TextField raw, Button 検索, Label status
        // 右: TableView + 件数 Label
        // 列: 依頼No, 希望納期, 調整納期, 製品, 原反, ユーザー, 入力日
        // 検索押下:
        //   criteria = new JuchuOrderSearchCriteria(from, to, product.getText(), raw.getText());
        //   optErr = criteria.validationError(); if present → status に表示し表クリア
        //   else → JuchuOrderSearch.filter(recordsSupplier.get(), criteria) を表にセット
    }
}
```

表示値の取り方:

- 依頼No / ユーザー: `OrderRecord` の getter
- 希望納期 / 調整納期 / 製品 / 入力日: `dbValues`
- 原反: `JuchuOrderSearch.displayRawMaterial(dbValues)`

JavaFX スレッド上でボタンハンドラを実行（通常の UI イベントで可）。バックグラウンドスレッドは不要（メモリ内フィルタ）。

- [ ] **Step 2: コンパイル確認**

```bat
cd code_java
.\mvnw.cmd -q -DskipTests compile
```

Expected: SUCCESS

- [ ] **Step 3: Commit**

```bash
git add code_java/src/main/java/jp/co/pm/ai/desktop/reconciliation/JuchuOrderSearchPane.java
git commit -m "feat: 受注検索タブ用の左条件・右結果 UI を追加"
```

---

### Task 4: ReconciliationApp 配線

**Files:**
- Modify: `code_java/src/main/java/jp/co/pm/ai/desktop/reconciliation/ReconciliationApp.java`（`buildEmbeddedRoot` 付近のタブ組み立て、約 931–954 行）

- [ ] **Step 1: タブ作成メソッド追加**

```java
private Tab createJuchuOrderSearchTab() {
    Tab tab = new Tab("受注検索");
    tab.setClosable(false);
    tab.setContent(JuchuOrderSearchPane.build(() -> List.copyOf(orderRecords)));
    return tab;
}
```

- [ ] **Step 2: タブ並びへ挿入**

`buildEmbeddedRoot` 内:

```java
Tab tabIndexSheet = createIndexSheetCatalogTab();
Tab tabJuchuSearch = createJuchuOrderSearchTab();
Tab tabSettings = createSettingsTab();
// ...
tabPane.getTabs()
        .addAll(
                tabVerification,
                tabIndexSheet,
                tabJuchuSearch,
                tabSettings,
                tabPostProcMaster,
                tabMasterList);
```

（既存の `createSettingsTab()` 呼び出し順が前後していても、最終 `addAll` 順が仕様どおりならよい。）

- [ ] **Step 3: コンパイル + 関連テスト**

```bat
cd code_java
.\mvnw.cmd -q test -Dtest=JuchuOrderSearchTest,MainShellInnerTabCatalogRequestFormInputTest
```

Expected: SUCCESS

- [ ] **Step 4: Commit**

```bash
git add code_java/src/main/java/jp/co/pm/ai/desktop/reconciliation/ReconciliationApp.java
git commit -m "feat: 依頼書入力に受注検索タブを追加"
```

---

### Task 5: 手動確認チェックリスト

- [ ] アプリ起動 → 依頼書入力 → 「受注検索」タブが表示される（目次シートの右）
- [ ] 納期未指定で検索 → エラーメッセージ、表は空
- [ ] 製品・原反両方空で検索 → エラーメッセージ
- [ ] 有効条件で件数と列が表示される
- [ ] タブ整理ツリーに「受注検索」が出る（`MainShellInnerTabCatalog`）
- [ ] マスター一覧の内側タブ（機械コード等）がタブ整理で壊れない

手動確認後、未コミットが残っていればまとめてコミット。push はユーザー指示時のみ。

---

## Spec coverage（自己レビュー）

| 仕様 | タスク |
|------|--------|
| 閲覧のみ・編集ジャンプなし | Task 3/4（遷移なし） |
| 受注行のみ | Task 4（`orderRecords`） |
| 希望／調整 OR・範囲必須 | Task 1 |
| 製品／原反 OR・片方必須 | Task 1 |
| 左条件・右結果 | Task 3 |
| タブ名・並び・カタログ | Task 2/4 |
| nested index ずれ | Task 2 |
| 単体テスト | Task 1/2 |
| 対象外（原本横断等） | 計画に含めず |

Placeholder / TBD: なし。型名は全タスクで `JuchuOrderSearch*` / `OrderRecord` で統一。
