# バージョン管理タブ（ダウングレード適用） Implementation Plan

> **For implementers:** Execute this plan task-by-task. Use the `test-driven-development` skill for each task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** メインシェルに「バージョン管理」タブを追加し、正本リリース直下＋`previous\v*` から選んだ版を既存ポータブル更新フローで強制適用する（ダウングレードは二重確認）。

**Architecture:** 一覧・正本解決は純粋な Java ヘルパ。適用は `PortableBundleSelfUpdateService` に「版比較スキップの強制適用」を追加し、タブから呼び出す。起動時自動更新の「新しいときだけ」は変更しない。

**Tech Stack:** Java 17+、JavaFX、JUnit 5、Maven（`code_java/mvnw.cmd`）

**仕様正本:** `docs/specs/2026-09-04-version-management-downgrade-design.md`

**前提:** パッケージ側の `previous\` 退避は `docs/plans/2026-09-04-package-release-generations.md` を先に完了推奨（一覧の実データ源）。ヘルパ／タブ自体はモックディレクトリでも単体テスト可能。

---

## ファイル構成

| ファイル | 責任 |
|----------|------|
| `.../config/PortableReleaseGenerationCatalog.java` | ZIP→親解決、世代一覧、`ReleaseGenerationRow` |
| `.../config/PortableReleaseGenerationCatalogTest.java` | 単体テスト |
| `.../config/PortableBundleSelfUpdateService.java` | `applyPortableBundleFromCanonical(...)` 強制適用 |
| `MainShellTabId.java` / `MainShellTabLayoutDefaults.java` / `MainShell.fxml` / `MainShellController.java` | タブ登録 |
| `fxml/VersionManagementTab.fxml` / `VersionManagementTabController.java` | UI |
| `VersionManagementTabFxmlTest.java` | FXML 構造テスト |

---

### Task 1: 正本リリース直下の解決 ＋ 世代一覧

**Files:**
- Create: `code_java/src/test/java/jp/co/pm/ai/desktop/config/PortableReleaseGenerationCatalogTest.java`
- Create: `code_java/src/main/java/jp/co/pm/ai/desktop/config/PortableReleaseGenerationCatalog.java`

- [ ] **Step 1: Write the failing test**

```java
package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class PortableReleaseGenerationCatalogTest {

    @TempDir Path tmp;

    @Test
    void resolveReleaseRoot_zipUsesParent() throws Exception {
        Path root = tmp.resolve("release");
        Files.createDirectories(root);
        Path zip = root.resolve("PMD_version_upgrade.zip");
        Files.writeString(zip, "x", StandardCharsets.UTF_8);
        assertEquals(root.toAbsolutePath().normalize(),
                PortableReleaseGenerationCatalog.resolveReleaseRoot(zip).orElseThrow());
    }

    @Test
    void resolveReleaseRoot_directoryIsItself() throws Exception {
        Path root = tmp.resolve("release");
        Files.createDirectories(root);
        assertEquals(root.toAbsolutePath().normalize(),
                PortableReleaseGenerationCatalog.resolveReleaseRoot(root).orElseThrow());
    }

    @Test
    void listGenerations_includesCurrentAndPrevious() throws Exception {
        Path root = tmp.resolve("release");
        Files.createDirectories(root);
        Files.writeString(root.resolve("version.txt"), "1.30.0\n", StandardCharsets.UTF_8);
        Files.writeString(root.resolve(PortableBundleSelfUpdater.PORTABLE_UPGRADE_ZIP_NAME), "u",
                StandardCharsets.UTF_8);
        Path prev = root.resolve("previous").resolve("v1.20.0");
        Files.createDirectories(prev);
        Files.writeString(prev.resolve("version.txt"), "1.20.0\n", StandardCharsets.UTF_8);
        Files.writeString(prev.resolve(PortableBundleSelfUpdater.PORTABLE_UPGRADE_ZIP_NAME), "u",
                StandardCharsets.UTF_8);

        List<PortableReleaseGenerationCatalog.ReleaseGenerationRow> rows =
                PortableReleaseGenerationCatalog.listGenerations(root);
        assertEquals(2, rows.size());
        assertTrue(rows.get(0).currentSlot());
        assertEquals(Optional.of(new BigDecimal("1.30.0")), rows.get(0).version());
        assertTrue(rows.get(0).hasUpgradeZip());
        assertFalse(rows.get(1).currentSlot());
        assertEquals(Optional.of(new BigDecimal("1.20.0")), rows.get(1).version());
        assertEquals(prev.toAbsolutePath().normalize(), rows.get(1).canonicalDir());
    }
}
```

- [ ] **Step 2: Run test to verify it fails**

```bat
cd code_java
mvnw.cmd -q test -Dtest=PortableReleaseGenerationCatalogTest
```

Expected: FAIL（クラス未存在）

- [ ] **Step 3: Write minimal implementation**

`PortableReleaseGenerationCatalog.java`:

```java
package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.Stream;

/** 正本リリース直下＋ previous\\v* の世代一覧（バージョン管理タブ用）。 */
public final class PortableReleaseGenerationCatalog {

    private PortableReleaseGenerationCatalog() {}

    public record ReleaseGenerationRow(
            Path canonicalDir,
            Optional<BigDecimal> version,
            boolean currentSlot,
            boolean hasUpgradeZip,
            String displayLabel) {}

    /** ZIP なら親、ディレクトリならそのパス。存在しない／その他は empty。 */
    public static Optional<Path> resolveReleaseRoot(Path sourceDirOrZip) {
        if (sourceDirOrZip == null) {
            return Optional.empty();
        }
        Path abs = sourceDirOrZip.toAbsolutePath().normalize();
        if (!Files.exists(abs)) {
            return Optional.empty();
        }
        if (Files.isRegularFile(abs)) {
            String name = abs.getFileName().toString().toLowerCase(Locale.ROOT);
            if (name.endsWith(".zip")) {
                Path parent = abs.getParent();
                return parent == null ? Optional.empty() : Optional.of(parent.normalize());
            }
            return Optional.empty();
        }
        if (Files.isDirectory(abs)) {
            return Optional.of(abs);
        }
        return Optional.empty();
    }

    public static List<ReleaseGenerationRow> listGenerations(Path releaseRoot) {
        Objects.requireNonNull(releaseRoot, "releaseRoot");
        Path root = releaseRoot.toAbsolutePath().normalize();
        List<ReleaseGenerationRow> out = new ArrayList<>();
        out.add(rowFor(root, true, "最新（リリース直下）"));
        Path prev = root.resolve("previous");
        if (Files.isDirectory(prev)) {
            try (Stream<Path> stream = Files.list(prev)) {
                stream.filter(Files::isDirectory)
                        .sorted(Comparator.comparing(PortableReleaseGenerationCatalog::mtime).reversed())
                        .forEach(dir -> out.add(rowFor(dir, false, dir.getFileName().toString())));
            } catch (IOException ignored) {
                // empty previous listing
            }
        }
        return List.copyOf(out);
    }

    private static long mtime(Path p) {
        try {
            return Files.getLastModifiedTime(p).toMillis();
        } catch (IOException e) {
            return 0L;
        }
    }

    private static ReleaseGenerationRow rowFor(Path dir, boolean currentSlot, String label) {
        Path abs = dir.toAbsolutePath().normalize();
        Optional<BigDecimal> ver =
                PortableBundleSelfUpdater.readCanonicalPortableBundleVersion(abs);
        boolean zip =
                Files.isRegularFile(abs.resolve(PortableBundleSelfUpdater.PORTABLE_UPGRADE_ZIP_NAME));
        return new ReleaseGenerationRow(abs, ver, currentSlot, zip, label);
    }
}
```

- [ ] **Step 4: Run test to verify it passes**

```bat
cd code_java
mvnw.cmd -q test -Dtest=PortableReleaseGenerationCatalogTest
```

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add code_java/src/main/java/jp/co/pm/ai/desktop/config/PortableReleaseGenerationCatalog.java \
  code_java/src/test/java/jp/co/pm/ai/desktop/config/PortableReleaseGenerationCatalogTest.java
git commit -m "$(cat <<'EOF'
feat: ポータブルリリース世代一覧カタログを追加

EOF
)"
```

---

### Task 2: 強制適用 API（版比較スキップ）

**Files:**
- Modify: `code_java/src/main/java/jp/co/pm/ai/desktop/config/PortableBundleSelfUpdateService.java`
- Test: 既存起動フローを壊さないことの確認は手動／既存テスト。新規は確認ダイアログを Service 外に置くため、Service は「確認済み前提で ZIP 適用開始」の public メソッドを追加。

- [ ] **Step 1: Extract / add public force-apply entry**

`PortableBundleSelfUpdateService` に追加（既存 `runZipUpgradeTask` を再利用）:

```java
/**
 * 版比較を行わず、指定正本（フォルダまたは ZIP）からアップグレード ZIP を適用する。
 * 確認ダイアログは呼び出し側で表示済みであること。
 *
 * @return 適用タスクを開始したら true（ZIP 無し・非ポータブルは false）
 */
public static boolean applyPortableBundleFromCanonical(
        PortableBundleProfile profile,
        Path canonicalSource,
        Map<String, String> ui,
        Stage dialogOwner,
        Consumer<String> log) {
    Objects.requireNonNull(profile, "profile");
    Objects.requireNonNull(canonicalSource, "canonicalSource");
    Path cwd = Path.of(System.getProperty("user.dir", ".")).toAbsolutePath().normalize();
    if (!profile.isPortableBundleLayout(cwd)) {
        logLine(log, "[version-mgmt] 強制適用不可（ポータブル配布レイアウトではありません）。");
        return false;
    }
    Path canonical = canonicalSource.toAbsolutePath().normalize();
    if (!PortableBundleSelfUpdater.isValidPortableBundleCanonical(canonical)) {
        logLine(log, "[version-mgmt] 正本が無効: " + PortableBundleSelfUpdater.safePathForLog(canonical));
        return false;
    }
    Optional<Path> upgradeZip = PortableBundleSelfUpdater.resolveEffectiveUpgradeZip(profile, canonical);
    if (upgradeZip.isEmpty()) {
        logLine(log, "[version-mgmt] アップグレード ZIP が見つかりません。");
        return false;
    }
    Optional<BigDecimal> cv =
            PortableBundleSelfUpdater.readCanonicalPortableBundleVersion(profile, canonical);
    String canonVerStr = cv.map(BigDecimal::toPlainString).orElse("?");
    runZipUpgradeTask(profile, cwd, canonical, upgradeZip.get(), canonVerStr, dialogOwner, log);
    return true;
}
```

`runZipUpgradeTask` が `private` のままなら、このメソッドから呼べる位置に置く（同クラス内）。

- [ ] **Step 2: Compile check**

```bat
cd code_java
mvnw.cmd -q compile
```

Expected: BUILD SUCCESS

- [ ] **Step 3: Commit**

```bash
git add code_java/src/main/java/jp/co/pm/ai/desktop/config/PortableBundleSelfUpdateService.java
git commit -m "$(cat <<'EOF'
feat: ポータブル配布の強制適用入口を追加

EOF
)"
```

---

### Task 3: MainShell タブ ID・レイアウト・FXML 骨格

**Files:**
- Modify: `MainShellTabId.java`（`TAB_ORGANIZER` の直前に追加）
- Modify: `MainShellTabLayoutDefaults.java`（`DEFAULT_FLAT` の `GLOBAL_SETTINGS` 付近、および `groupedLayout` の「環境設定」）
- Create: `VersionManagementTab.fxml` / `VersionManagementTabController.java`
- Modify: `MainShell.fxml` / `MainShellController.java`
- Create: `VersionManagementTabFxmlTest.java`

- [ ] **Step 1: Write FXML structure test (failing)**

```java
package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Map;
import java.util.stream.IntStream;

import javax.xml.parsers.DocumentBuilderFactory;

import org.junit.jupiter.api.Test;
import org.w3c.dom.Element;

class VersionManagementTabFxmlTest {

    @Test
    void applyButtonAndTablePresent() throws Exception {
        var resource =
                VersionManagementTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/VersionManagementTab.fxml");
        assertTrue(resource != null);
        var document = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(resource);
        var buttons = document.getElementsByTagName("Button");
        Map<String, Element> byId =
                IntStream.range(0, buttons.getLength())
                        .mapToObj(buttons::item)
                        .filter(Element.class::isInstance)
                        .map(Element.class::cast)
                        .filter(b -> b.hasAttribute("fx:id"))
                        .collect(
                                java.util.stream.Collectors.toMap(
                                        b -> b.getAttribute("fx:id"), b -> b));
        assertTrue(byId.containsKey("refreshButton"));
        assertTrue(byId.containsKey("applyButton"));
        assertEquals("この版を適用", byId.get("applyButton").getAttribute("text"));
        assertEquals("#onApplyButtonAction", byId.get("applyButton").getAttribute("onAction"));
    }
}
```

- [ ] **Step 2: Run to verify fail**

```bat
mvnw.cmd -q test -Dtest=VersionManagementTabFxmlTest
```

Expected: FAIL（FXML 無し）

- [ ] **Step 3: Add enum + layout + FXML + controller stub + MainShell wire**

`MainShellTabId`（`TAB_ORGANIZER` 直前）:

```java
    /** ポータブル配布の版一覧・適用（アップ／ダウングレード）。 */
    VERSION_MANAGEMENT("versionManagement"),
```

`DEFAULT_FLAT_TAB_KEY_ORDER`: `GLOBAL_SETTINGS.key()` の直後に `VERSION_MANAGEMENT.key()` を挿入。

`groupedLayout` の「環境設定」リストに `VERSION_MANAGEMENT` を `GLOBAL_SETTINGS` の次へ追加。

`VersionManagementTab.fxml`（最小）:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<?import javafx.geometry.Insets?>
<?import javafx.scene.control.*?>
<?import javafx.scene.layout.*?>
<VBox xmlns="http://javafx.com/javafx/21" xmlns:fx="http://javafx.com/fxml/1"
      fx:controller="jp.co.pm.ai.desktop.VersionManagementTabController"
      spacing="8">
    <padding><Insets top="12" right="12" bottom="12" left="12"/></padding>
    <Label fx:id="statusLabel" text="正本パスを解決しています…"/>
    <Label fx:id="localVersionLabel" text="ローカル版: —"/>
    <HBox spacing="8">
        <Button fx:id="refreshButton" text="再読込" onAction="#onRefreshButtonAction"/>
        <Button fx:id="applyButton" text="この版を適用" onAction="#onApplyButtonAction"/>
    </HBox>
    <TableView fx:id="generationTable" VBox.vgrow="ALWAYS"/>
</VBox>
```

`VersionManagementTabController.java`（骨格: `bindShell`・空の `onRefresh` / `onApply`）:

```java
package jp.co.pm.ai.desktop;

import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TableView;

public class VersionManagementTabController {

    @FXML private Label statusLabel;
    @FXML private Label localVersionLabel;
    @FXML private Button refreshButton;
    @FXML private Button applyButton;
    @FXML private TableView<Object> generationTable;

    private MainShellController shell;

    public void bindShell(MainShellController shell) {
        this.shell = shell;
    }

    @FXML
    private void onRefreshButtonAction() {
        // Task 4
    }

    @FXML
    private void onApplyButtonAction() {
        // Task 4
    }
}
```

`MainShell.fxml`: `mainShellTabGlobalSettings` の直後に:

```xml
                <Tab fx:id="mainShellTabVersionManagement" closable="false" text="バージョン管理">
                    <content>
                        <fx:include fx:id="versionManagementTab" source="VersionManagementTab.fxml"/>
                    </content>
                </Tab>
```

`MainShellController`: `@FXML Tab mainShellTabVersionManagement`、`@FXML VersionManagementTabController versionManagementTabController`、`bindShell`、`mainShellTabFor` の `case VERSION_MANAGEMENT -> mainShellTabVersionManagement`。既存の `fx:id` include 命名規則に合わせ、必要なら `versionManagementTabController` の注入名を FXMLLoader 規約どおりにする。

- [ ] **Step 4: Run FXML test**

```bat
mvnw.cmd -q test -Dtest=VersionManagementTabFxmlTest
```

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add code_java/src/main/java/jp/co/pm/ai/desktop/MainShellTabId.java \
  code_java/src/main/java/jp/co/pm/ai/desktop/config/MainShellTabLayoutDefaults.java \
  code_java/src/main/resources/jp/co/pm/ai/desktop/fxml/MainShell.fxml \
  code_java/src/main/resources/jp/co/pm/ai/desktop/fxml/VersionManagementTab.fxml \
  code_java/src/main/java/jp/co/pm/ai/desktop/VersionManagementTabController.java \
  code_java/src/main/java/jp/co/pm/ai/desktop/MainShellController.java \
  code_java/src/test/java/jp/co/pm/ai/desktop/VersionManagementTabFxmlTest.java
git commit -m "$(cat <<'EOF'
feat: バージョン管理タブをメインシェルに追加

EOF
)"
```

---

### Task 4: 一覧表示・適用確認（ダウングレード二重確認）

**Files:**
- Modify: `VersionManagementTabController.java`

- [ ] **Step 1: Implement refresh**

- `shell.snapshotUiEnv()` から `AppPaths.KEY_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR` を読む
- `resolveReleaseRoot` → `listGenerations`
- ローカル版は `PortableBundleSelfUpdater.readLocalBundleVersion`（ポータブルなら cwd + pm-ai-data）
- `TableView` の列: 表示名、版、Upgrade ZIP 有無、現在スロット
- 行型は `ReleaseGenerationRow` に変更（Task 3 の `Object` を置き換え）

- [ ] **Step 2: Implement apply with confirms**

```text
1. 選択行なし → return
2. !hasUpgradeZip → Alert 警告
3. 非ポータブル → Alert で適用不可
4. 選択版とローカルを BigDecimal 比較
   - 選択 > ローカル: CONFIRMATION 1 回（アップ）
   - 選択 <= ローカル: CONFIRMATION 1 回目（ダウングレード警告）→ OK なら 2 回目「本当にこの版に戻しますか」
5. OK 後: PortableBundleSelfUpdateService.applyPortableBundleFromCanonical(
     PortableBundleProfile.PMD, row.canonicalDir(), shell.snapshotUiEnv(), shell の Stage, shell::appendLog)
```

ダウングレード判定: `selected.version()` が empty なら「不明版」として二重確認扱い。

- [ ] **Step 3: Manual smoke（ポータブル環境がある場合）**

1. タブを開き一覧に最新＋ previous が出る
2. 新しい版: 確認1回
3. 古い版: 確認2回

- [ ] **Step 4: Commit**

```bash
git add code_java/src/main/java/jp/co/pm/ai/desktop/VersionManagementTabController.java
git commit -m "$(cat <<'EOF'
feat: バージョン管理タブで世代適用とダウングレード二重確認

EOF
)"
```

---

## Spec coverage（本プラン）

| 要件 | Task |
|------|------|
| ZIP→親 / フォルダ直下 | Task 1 |
| 最新＋previous 一覧 | Task 1, 4 |
| versionManagement タブ・環境設定グループ | Task 3 |
| 強制適用（起動時ロジック非変更） | Task 2 |
| ダウン二重確認 / アップ1回 | Task 4 |
| 非ポータブルは適用不可 | Task 4 |

パッケージ退避は前プラン。
