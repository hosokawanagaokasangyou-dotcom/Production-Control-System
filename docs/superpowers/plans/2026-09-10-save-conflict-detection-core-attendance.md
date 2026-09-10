# 保存前ファイル競合検知（共通基盤＋勤怠系）Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 明示保存の直前に SHA-256 で外部変更を検知し、業務語要約付きダイアログ（再読込して更新／強制上書き／キャンセル）を出す共通基盤を入れ、会社カレンダー・メンバー勤怠・機械カレンダーに適用する。

**Architecture:** `FileContentFingerprint` でパスごとのハッシュ／欠落 sentinel を記録し、`SaveConflictChecker` が保存前に再計算して一致判定する。不一致時は画面別 `ConflictDiffSummarizer` が業務語要約を作り、`ConflictSaveDialog`（強調ヘッダ＋スクロール）でユーザー選択を返す。常時 OS ロックは入れない。

**Tech Stack:** Java 21+ / JavaFX / JUnit 5 / Maven（`code_java/mvnw.cmd`）

**仕様正本:** `docs/superpowers/specs/2026-09-10-save-conflict-detection-design.md`

**後続計画:**
- `docs/superpowers/plans/2026-09-10-save-conflict-detection-dispatch-screens.md`
- `docs/superpowers/plans/2026-09-10-save-conflict-detection-request-and-settings.md`

---

## File map

| ファイル | 役割 |
|----------|------|
| Create: `code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/FileContentFingerprint.java` | SHA-256／ABSENT |
| Create: `code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/FingerprintBaseline.java` | 複数パスのハッシュ＋スナップショットバイト |
| Create: `code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/SaveConflictChecker.java` | baseline とディスク比較 |
| Create: `code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/ConflictCheckResult.java` | OK / CONFLICT / IO_ERROR |
| Create: `code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/ConflictDiffSummarizer.java` | 要約 interface |
| Create: `code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/ConflictSaveChoice.java` | RELOAD / OVERWRITE / CANCEL |
| Create: `code_java/src/main/java/jp/co/pm/ai/desktop/ui/ConflictSaveDialog.java` | 強調ヘッダ＋スクロール要約ダイアログ |
| Create: `code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/AttendanceConflictDiffSummarizer.java` | 勤怠 JSON／xlsx 向け業務語要約 |
| Create: `code_java/src/test/java/jp/co/pm/ai/desktop/io/conflict/FileContentFingerprintTest.java` | 指紋単体 |
| Create: `code_java/src/test/java/jp/co/pm/ai/desktop/io/conflict/SaveConflictCheckerTest.java` | 判定単体 |
| Create: `code_java/src/test/java/jp/co/pm/ai/desktop/io/conflict/AttendanceConflictDiffSummarizerTest.java` | 勤怠要約単体 |
| Modify: `MemberAttendanceTabController.java` | baseline 保持・保存前ガード |
| Modify: `CompanyCalendarTabController.java` | 同上 |
| Modify: `MachineCalendarTabController.java` | 同上 |

---

### Task 1: FileContentFingerprint（TDD）

**Files:**
- Create: `code_java/src/test/java/jp/co/pm/ai/desktop/io/conflict/FileContentFingerprintTest.java`
- Create: `code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/FileContentFingerprint.java`

- [ ] **Step 1: 失敗するテストを書く**

```java
package jp.co.pm.ai.desktop.io.conflict;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class FileContentFingerprintTest {

    @TempDir Path dir;

    @Test
    void absentPath_returnsAbsentSentinel() throws Exception {
        Path missing = dir.resolve("no-such.json");
        assertEquals(FileContentFingerprint.ABSENT, FileContentFingerprint.sha256Hex(missing));
    }

    @Test
    void sameBytes_sameHash() throws Exception {
        Path f = dir.resolve("a.json");
        Files.writeString(f, "{\"x\":1}", StandardCharsets.UTF_8);
        String h1 = FileContentFingerprint.sha256Hex(f);
        String h2 = FileContentFingerprint.sha256Hex(f);
        assertEquals(h1, h2);
        assertTrue(h1.matches("[0-9a-f]{64}"));
    }

    @Test
    void changedBytes_differentHash() throws Exception {
        Path f = dir.resolve("a.json");
        Files.writeString(f, "{\"x\":1}", StandardCharsets.UTF_8);
        String h1 = FileContentFingerprint.sha256Hex(f);
        Files.writeString(f, "{\"x\":2}", StandardCharsets.UTF_8);
        assertNotEquals(h1, FileContentFingerprint.sha256Hex(f));
    }

    @Test
    void hashOfBytes_matchesFile() throws Exception {
        byte[] bytes = "{\"y\":3}".getBytes(StandardCharsets.UTF_8);
        Path f = dir.resolve("b.json");
        Files.write(f, bytes);
        assertEquals(FileContentFingerprint.sha256Hex(bytes), FileContentFingerprint.sha256Hex(f));
    }
}
```

- [ ] **Step 2: テスト実行（失敗確認）**

Run: `cd code_java && .\mvnw.cmd -q test "-Dtest=FileContentFingerprintTest"`

Expected: FAIL（クラス未定義）

- [ ] **Step 3: 最小実装**

```java
package jp.co.pm.ai.desktop.io.conflict;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.HexFormat;

public final class FileContentFingerprint {

    /** ファイルが存在しないときの sentinel（64 桁 hex と衝突しない固定文字列）。 */
    public static final String ABSENT = "ABSENT";

    private FileContentFingerprint() {}

    public static String sha256Hex(Path path) throws IOException {
        if (path == null || !Files.isRegularFile(path)) {
            return ABSENT;
        }
        return sha256Hex(Files.readAllBytes(path));
    }

    public static String sha256Hex(byte[] bytes) {
        if (bytes == null) {
            throw new IllegalArgumentException("bytes");
        }
        try {
            MessageDigest md = MessageDigest.getInstance("SHA-256");
            return HexFormat.of().formatHex(md.digest(bytes));
        } catch (NoSuchAlgorithmException e) {
            throw new IllegalStateException("SHA-256 unavailable", e);
        }
    }
}
```

- [ ] **Step 4: テスト再実行（成功確認）**

Run: `cd code_java && .\mvnw.cmd -q test "-Dtest=FileContentFingerprintTest"`

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/FileContentFingerprint.java \
  code_java/src/test/java/jp/co/pm/ai/desktop/io/conflict/FileContentFingerprintTest.java
git commit -m "feat: ファイル内容 SHA-256 指紋ユーティリティを追加"
```

---

### Task 2: FingerprintBaseline + SaveConflictChecker（TDD）

**Files:**
- Create: `code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/FingerprintBaseline.java`
- Create: `code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/ConflictCheckResult.java`
- Create: `code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/SaveConflictChecker.java`
- Create: `code_java/src/test/java/jp/co/pm/ai/desktop/io/conflict/SaveConflictCheckerTest.java`

- [ ] **Step 1: 失敗するテストを書く**

```java
package jp.co.pm.ai.desktop.io.conflict;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class SaveConflictCheckerTest {

    @TempDir Path dir;

    @Test
    void unchangedSet_isOk() throws Exception {
        Path json = dir.resolve("attendance-data.json");
        Path xlsx = dir.resolve("勤怠・機械カレンダー.xlsx");
        Files.writeString(json, "{\"members\":[]}", StandardCharsets.UTF_8);
        Files.write(xlsx, new byte[] {1, 2, 3});
        FingerprintBaseline baseline = FingerprintBaseline.capture(List.of(json, xlsx));
        ConflictCheckResult r = SaveConflictChecker.check(baseline);
        assertEquals(ConflictCheckResult.Kind.OK, r.kind());
    }

    @Test
    void changedRelatedFile_isConflict() throws Exception {
        Path json = dir.resolve("attendance-data.json");
        Path xlsx = dir.resolve("勤怠・機械カレンダー.xlsx");
        Files.writeString(json, "{\"members\":[]}", StandardCharsets.UTF_8);
        Files.write(xlsx, new byte[] {1, 2, 3});
        FingerprintBaseline baseline = FingerprintBaseline.capture(List.of(json, xlsx));
        Files.write(xlsx, new byte[] {9, 9, 9});
        ConflictCheckResult r = SaveConflictChecker.check(baseline);
        assertEquals(ConflictCheckResult.Kind.CONFLICT, r.kind());
        assertTrue(r.mismatchedPaths().contains(xlsx));
    }

    @Test
    void absentThenAppears_isConflict() throws Exception {
        Path json = dir.resolve("attendance-data.json");
        FingerprintBaseline baseline = FingerprintBaseline.capture(List.of(json));
        Files.writeString(json, "{}", StandardCharsets.UTF_8);
        assertEquals(ConflictCheckResult.Kind.CONFLICT, SaveConflictChecker.check(baseline).kind());
    }

    @Test
    void presentThenMissing_isConflict() throws Exception {
        Path json = dir.resolve("attendance-data.json");
        Files.writeString(json, "{}", StandardCharsets.UTF_8);
        FingerprintBaseline baseline = FingerprintBaseline.capture(List.of(json));
        Files.delete(json);
        assertEquals(ConflictCheckResult.Kind.CONFLICT, SaveConflictChecker.check(baseline).kind());
    }
}
```

- [ ] **Step 2: テスト実行（失敗確認）**

Run: `cd code_java && .\mvnw.cmd -q test "-Dtest=SaveConflictCheckerTest"`

Expected: FAIL

- [ ] **Step 3: 最小実装**

`ConflictCheckResult.java`:

```java
package jp.co.pm.ai.desktop.io.conflict;

import java.nio.file.Path;
import java.util.List;
import java.util.Objects;

public final class ConflictCheckResult {

    public enum Kind {
        OK,
        CONFLICT,
        IO_ERROR
    }

    private final Kind kind;
    private final List<Path> mismatchedPaths;
    private final String errorMessage;

    private ConflictCheckResult(Kind kind, List<Path> mismatchedPaths, String errorMessage) {
        this.kind = Objects.requireNonNull(kind);
        this.mismatchedPaths = mismatchedPaths == null ? List.of() : List.copyOf(mismatchedPaths);
        this.errorMessage = errorMessage;
    }

    public static ConflictCheckResult ok() {
        return new ConflictCheckResult(Kind.OK, List.of(), null);
    }

    public static ConflictCheckResult conflict(List<Path> mismatched) {
        return new ConflictCheckResult(Kind.CONFLICT, mismatched, null);
    }

    public static ConflictCheckResult ioError(String message) {
        return new ConflictCheckResult(Kind.IO_ERROR, List.of(), message);
    }

    public Kind kind() {
        return kind;
    }

    public List<Path> mismatchedPaths() {
        return mismatchedPaths;
    }

    public String errorMessage() {
        return errorMessage;
    }
}
```

`FingerprintBaseline.java`:

```java
package jp.co.pm.ai.desktop.io.conflict;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

public final class FingerprintBaseline {

    private final Map<Path, String> hashes;
    private final Map<Path, byte[]> snapshots;

    private FingerprintBaseline(Map<Path, String> hashes, Map<Path, byte[]> snapshots) {
        this.hashes = Collections.unmodifiableMap(new LinkedHashMap<>(hashes));
        this.snapshots = Collections.unmodifiableMap(new LinkedHashMap<>(snapshots));
    }

    public static FingerprintBaseline capture(List<Path> paths) throws IOException {
        Objects.requireNonNull(paths, "paths");
        Map<Path, String> hashes = new LinkedHashMap<>();
        Map<Path, byte[]> snapshots = new LinkedHashMap<>();
        for (Path p : paths) {
            Path abs = p.toAbsolutePath().normalize();
            if (Files.isRegularFile(abs)) {
                byte[] bytes = Files.readAllBytes(abs);
                snapshots.put(abs, bytes);
                hashes.put(abs, FileContentFingerprint.sha256Hex(bytes));
            } else {
                snapshots.put(abs, new byte[0]);
                hashes.put(abs, FileContentFingerprint.ABSENT);
            }
        }
        return new FingerprintBaseline(hashes, snapshots);
    }

    public Map<Path, String> hashes() {
        return hashes;
    }

    public Map<Path, byte[]> snapshots() {
        return snapshots;
    }

    public List<Path> paths() {
        return List.copyOf(hashes.keySet());
    }
}
```

`SaveConflictChecker.java`:

```java
package jp.co.pm.ai.desktop.io.conflict;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

public final class SaveConflictChecker {

    private SaveConflictChecker() {}

    public static ConflictCheckResult check(FingerprintBaseline baseline) {
        Objects.requireNonNull(baseline, "baseline");
        List<Path> mismatched = new ArrayList<>();
        try {
            for (Path path : baseline.paths()) {
                String current = FileContentFingerprint.sha256Hex(path);
                String expected = baseline.hashes().get(path);
                if (!Objects.equals(current, expected)) {
                    mismatched.add(path);
                }
            }
        } catch (IOException e) {
            return ConflictCheckResult.ioError(e.getMessage() != null ? e.getMessage() : e.toString());
        }
        if (mismatched.isEmpty()) {
            return ConflictCheckResult.ok();
        }
        return ConflictCheckResult.conflict(mismatched);
    }
}
```

- [ ] **Step 4: テスト再実行**

Run: `cd code_java && .\mvnw.cmd -q test "-Dtest=SaveConflictCheckerTest"`

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/FingerprintBaseline.java \
  code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/ConflictCheckResult.java \
  code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/SaveConflictChecker.java \
  code_java/src/test/java/jp/co/pm/ai/desktop/io/conflict/SaveConflictCheckerTest.java
git commit -m "feat: 保存前の指紋セット競合判定を追加"
```

---

### Task 3: ConflictDiffSummarizer インタフェース + ConflictSaveChoice + Dialog

**Files:**
- Create: `code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/ConflictDiffSummarizer.java`
- Create: `code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/ConflictSaveChoice.java`
- Create: `code_java/src/main/java/jp/co/pm/ai/desktop/ui/ConflictSaveDialog.java`
- Create: `code_java/src/test/java/jp/co/pm/ai/desktop/io/conflict/ConflictSaveChoiceTest.java`

- [ ] **Step 1: Choice の単純テスト**

```java
package jp.co.pm.ai.desktop.io.conflict;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class ConflictSaveChoiceTest {

    @Test
    void values_areStable() {
        assertEquals(3, ConflictSaveChoice.values().length);
        assertEquals(ConflictSaveChoice.RELOAD, ConflictSaveChoice.valueOf("RELOAD"));
        assertEquals(ConflictSaveChoice.OVERWRITE, ConflictSaveChoice.valueOf("OVERWRITE"));
        assertEquals(ConflictSaveChoice.CANCEL, ConflictSaveChoice.valueOf("CANCEL"));
    }
}
```

- [ ] **Step 2: 実装**

`ConflictSaveChoice.java`:

```java
package jp.co.pm.ai.desktop.io.conflict;

public enum ConflictSaveChoice {
    RELOAD,
    OVERWRITE,
    CANCEL
}
```

`ConflictDiffSummarizer.java`:

```java
package jp.co.pm.ai.desktop.io.conflict;

import java.nio.file.Path;
import java.util.Map;

@FunctionalInterface
public interface ConflictDiffSummarizer {

    /**
     * @param baselineSnapshots 読込時バイト（欠落は空配列）
     * @param diskBytes 現ディスクバイト（欠落は空配列、キーは absolute normalize）
     * @param mismatched 不一致パス
     * @return 業務語の要約（改行可）。失敗時はフォールバック文言でもよい
     */
    String summarize(
            Map<Path, byte[]> baselineSnapshots,
            Map<Path, byte[]> diskBytes,
            java.util.List<Path> mismatched);
}
```

`ConflictSaveDialog.java`（UI スレッドで呼ぶ）:

```java
package jp.co.pm.ai.desktop.ui;

import java.util.Optional;

import javafx.geometry.Insets;
import javafx.scene.control.ButtonBar;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;
import javafx.scene.layout.VBox;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.io.conflict.ConflictSaveChoice;

public final class ConflictSaveDialog {

    private ConflictSaveDialog() {}

    public static ConflictSaveChoice show(Window owner, String screenTitle, String summary) {
        Dialog<ConflictSaveChoice> dialog = new Dialog<>();
        if (owner != null) {
            dialog.initOwner(owner);
        }
        dialog.setTitle("外部変更の検出");
        dialog.setHeaderText("⚠ 競合 — " + (screenTitle != null ? screenTitle : "保存先") + " が他で変更されています");
        dialog.getDialogPane().setStyle("-fx-header-color: #5c2b2b;");

        Label caption = new Label("変更の要約");
        caption.getStyleClass().add("label");
        TextArea area = new TextArea(summary != null ? summary : "");
        area.setEditable(false);
        area.setWrapText(true);
        area.setPrefRowCount(12);
        area.setPrefWidth(560);
        Label note = new Label("「再読込して更新」を選ぶと、いまの未保存編集は破棄されます。");
        note.setWrapText(true);

        VBox body = new VBox(8, caption, area, note);
        body.setPadding(new Insets(8));
        dialog.getDialogPane().setContent(body);

        ButtonType reload = new ButtonType("再読込して更新", ButtonBar.ButtonData.LEFT);
        ButtonType overwrite = new ButtonType("強制上書き", ButtonBar.ButtonData.OTHER);
        ButtonType cancel = new ButtonType("キャンセル", ButtonBar.ButtonData.CANCEL_CLOSE);
        dialog.getDialogPane().getButtonTypes().setAll(reload, overwrite, cancel);
        dialog.setResultConverter(
                bt -> {
                    if (bt == reload) {
                        return ConflictSaveChoice.RELOAD;
                    }
                    if (bt == overwrite) {
                        return ConflictSaveChoice.OVERWRITE;
                    }
                    return ConflictSaveChoice.CANCEL;
                });
        Optional<ConflictSaveChoice> ans = dialog.showAndWait();
        return ans.orElse(ConflictSaveChoice.CANCEL);
    }
}
```

- [ ] **Step 3: テスト**

Run: `cd code_java && .\mvnw.cmd -q test "-Dtest=ConflictSaveChoiceTest"`

Expected: PASS（Dialog は手動確認。コンパイルに含める）

- [ ] **Step 4: Commit**

```bash
git add code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/ConflictDiffSummarizer.java \
  code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/ConflictSaveChoice.java \
  code_java/src/main/java/jp/co/pm/ai/desktop/ui/ConflictSaveDialog.java \
  code_java/src/test/java/jp/co/pm/ai/desktop/io/conflict/ConflictSaveChoiceTest.java
git commit -m "feat: 保存競合ダイアログと要約インタフェースを追加"
```

---

### Task 4: AttendanceConflictDiffSummarizer（TDD）

**Files:**
- Create: `code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/AttendanceConflictDiffSummarizer.java`
- Create: `code_java/src/test/java/jp/co/pm/ai/desktop/io/conflict/AttendanceConflictDiffSummarizerTest.java`

勤怠正本 JSON の代表構造（メンバー配列・日次セル）を前提に、氏名の増減とセル変更件数を要約する。xlsx はバイナリのため「関連 Excel が変更されています」と出す。

- [ ] **Step 1: 失敗するテスト**

```java
package jp.co.pm.ai.desktop.io.conflict;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

class AttendanceConflictDiffSummarizerTest {

    @Test
    void summarizesMemberAddedAndCellChanges() {
        Path json = Path.of("attendance-data.json").toAbsolutePath().normalize();
        Path xlsx = Path.of("勤怠・機械カレンダー.xlsx").toAbsolutePath().normalize();
        String base =
                """
                {"members":[{"id":"1","name":"佐藤"}],"member_attendance":{"1":{"2026-09-01":"出勤"}}}
                """;
        String disk =
                """
                {"members":[{"id":"1","name":"佐藤"},{"id":"2","name":"山田"}],"member_attendance":{"1":{"2026-09-01":"公休"},"2":{"2026-09-01":"出勤"}}}
                """;
        byte[] xBase = new byte[] {1};
        byte[] xDisk = new byte[] {2};
        String summary =
                new AttendanceConflictDiffSummarizer()
                        .summarize(
                                Map.of(json, base.getBytes(StandardCharsets.UTF_8), xlsx, xBase),
                                Map.of(json, disk.getBytes(StandardCharsets.UTF_8), xlsx, xDisk),
                                List.of(json, xlsx));
        assertTrue(summary.contains("山田"), summary);
        assertTrue(summary.contains("追加") || summary.contains("増"), summary);
        assertTrue(summary.contains("Excel") || summary.contains("xlsx") || summary.contains("カレンダー"), summary);
    }

    @Test
    void fallbackWhenJsonBroken() {
        Path json = Path.of("attendance-data.json").toAbsolutePath().normalize();
        String summary =
                new AttendanceConflictDiffSummarizer()
                        .summarize(
                                Map.of(json, "{".getBytes(StandardCharsets.UTF_8)),
                                Map.of(json, "}".getBytes(StandardCharsets.UTF_8)),
                                List.of(json));
        assertTrue(summary.contains("詳細差分を生成できませんでした") || summary.contains("attendance-data"), summary);
    }
}
```

- [ ] **Step 2: 失敗確認**

Run: `cd code_java && .\mvnw.cmd -q test "-Dtest=AttendanceConflictDiffSummarizerTest"`

- [ ] **Step 3: 実装**

`AttendanceConflictDiffSummarizer` は Jackson `ObjectMapper` で `members`（`name`）と `member_attendance` / `company_calendar`（会社カレンダー用）を比較する。実装時に実 JSON キーが異なる場合は、既存の勤怠読込コード（`MemberAttendanceTabController` / Python merge のパッチ形）に合わせてキー名を修正する。xlsx パスはファイル名に「カレンダー」または `.xlsx` を含むものを関連 Excel として扱う。

最低限の骨子:

```java
package jp.co.pm.ai.desktop.io.conflict;

import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

public final class AttendanceConflictDiffSummarizer implements ConflictDiffSummarizer {

    private static final ObjectMapper JSON = new ObjectMapper();

    @Override
    public String summarize(
            Map<Path, byte[]> baselineSnapshots,
            Map<Path, byte[]> diskBytes,
            List<Path> mismatched) {
        List<String> lines = new ArrayList<>();
        try {
            for (Path p : mismatched) {
                String name = p.getFileName().toString();
                if (name.endsWith(".xlsx") || name.endsWith(".xlsm")) {
                    lines.add("・関連 Excel（" + name + "）が変更されています");
                    continue;
                }
                JsonNode base = JSON.readTree(baselineSnapshots.getOrDefault(p, new byte[0]));
                JsonNode disk = JSON.readTree(diskBytes.getOrDefault(p, new byte[0]));
                lines.addAll(diffMembers(base, disk));
                lines.addAll(diffAttendanceCells(base, disk));
                lines.addAll(diffCompanyDays(base, disk));
            }
            if (lines.isEmpty()) {
                lines.add("・内容が変更されています（詳細キーは特定できませんでした）");
            }
            return String.join("\n", lines);
        } catch (Exception e) {
            StringBuilder sb = new StringBuilder("詳細差分を生成できませんでした。\n");
            for (Path p : mismatched) {
                String hex = FileContentFingerprint.sha256Hex(diskBytes.getOrDefault(p, new byte[0]));
                sb.append("・")
                        .append(p.getFileName())
                        .append(" hash=")
                        .append(hex, 0, Math.min(12, hex.length()))
                        .append("…\n");
            }
            return sb.toString().trim();
        }
    }

    private static Set<String> memberNames(JsonNode root) {
        Set<String> names = new LinkedHashSet<>();
        JsonNode members = root.path("members");
        if (members.isArray()) {
            for (JsonNode m : members) {
                String n = m.path("name").asText("").trim();
                if (!n.isEmpty()) {
                    names.add(n);
                }
            }
        }
        return names;
    }

    private static List<String> diffMembers(JsonNode base, JsonNode disk) {
        Set<String> b = memberNames(base);
        Set<String> d = memberNames(disk);
        List<String> lines = new ArrayList<>();
        for (String n : d) {
            if (!b.contains(n)) {
                lines.add("・メンバー「" + n + "」が追加されています");
            }
        }
        for (String n : b) {
            if (!d.contains(n)) {
                lines.add("・メンバー「" + n + "」が削除されています");
            }
        }
        return lines;
    }

    private static List<String> diffAttendanceCells(JsonNode base, JsonNode disk) {
        // member_attendance: { memberId: { yyyy-MM-dd: code } }
        int changed = 0;
        JsonNode b = base.path("member_attendance");
        JsonNode d = disk.path("member_attendance");
        if (d.isObject()) {
            var fields = d.fields();
            while (fields.hasNext()) {
                var e = fields.next();
                JsonNode days = e.getValue();
                if (!days.isObject()) {
                    continue;
                }
                var dayFields = days.fields();
                while (dayFields.hasNext()) {
                    var day = dayFields.next();
                    String was = b.path(e.getKey()).path(day.getKey()).asText("");
                    String now = day.getValue().asText("");
                    if (!was.equals(now)) {
                        changed++;
                    }
                }
            }
        }
        if (changed == 0) {
            return List.of();
        }
        return List.of("・メンバー勤怠セルが " + changed + " 件変更されています");
    }

    private static List<String> diffCompanyDays(JsonNode base, JsonNode disk) {
        // company_calendar 等: 日付→区分。キー名は実装時に実データへ合わせる
        JsonNode b = base.path("company_calendar");
        JsonNode d = disk.path("company_calendar");
        if (!d.isObject() && !b.isObject()) {
            return List.of();
        }
        int changed = 0;
        if (d.isObject()) {
            var it = d.fields();
            while (it.hasNext()) {
                var e = it.next();
                if (!b.path(e.getKey()).asText("").equals(e.getValue().asText(""))) {
                    changed++;
                }
            }
        }
        if (changed == 0) {
            return List.of();
        }
        return List.of("・会社カレンダーの日付区分が " + changed + " 件変更されています");
    }
}
```

実装後、実 `attendance-data.json` のキーが違う場合はテストのサンプル JSON と実装を揃える（`code/` 配下のサンプルまたは工場キャッシュを参照）。

- [ ] **Step 4: テスト成功**

Run: `cd code_java && .\mvnw.cmd -q test "-Dtest=AttendanceConflictDiffSummarizerTest"`

- [ ] **Step 5: Commit**

```bash
git add code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/AttendanceConflictDiffSummarizer.java \
  code_java/src/test/java/jp/co/pm/ai/desktop/io/conflict/AttendanceConflictDiffSummarizerTest.java
git commit -m "feat: 勤怠 JSON／Excel の業務語競合要約を追加"
```

---

### Task 5: 共通ヘルパ SaveConflictGate（コントローラから呼びやすくする）

**Files:**
- Create: `code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/SaveConflictGate.java`
- Create: `code_java/src/test/java/jp/co/pm/ai/desktop/io/conflict/SaveConflictGateTest.java`

ゲートは「ダイアログを出さない純粋判定」と「ディスクバイト読込」を提供する。UI ダイアログはコントローラ側で `ConflictSaveDialog` を呼ぶ（テスト容易性のため）。

```java
package jp.co.pm.ai.desktop.io.conflict;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.Map;

public final class SaveConflictGate {

    public enum Proceed {
        /** 競合なし、またはユーザーが強制上書きを選んだ想定で保存続行 */
        CONTINUE_SAVE,
        /** ユーザーが再読込を選んだ／呼び出し側で reload する */
        RELOAD_UI,
        /** キャンセルまたは I/O エラーで中止 */
        ABORT
    }

    private SaveConflictGate() {}

    public static Map<Path, byte[]> readDiskBytes(FingerprintBaseline baseline) throws IOException {
        Map<Path, byte[]> disk = new LinkedHashMap<>();
        for (Path p : baseline.paths()) {
            if (Files.isRegularFile(p)) {
                disk.put(p, Files.readAllBytes(p));
            } else {
                disk.put(p, new byte[0]);
            }
        }
        return disk;
    }

    public static String summarizeOrFallback(
            ConflictDiffSummarizer summarizer,
            FingerprintBaseline baseline,
            Map<Path, byte[]> diskBytes,
            ConflictCheckResult check) {
        try {
            return summarizer.summarize(baseline.snapshots(), diskBytes, check.mismatchedPaths());
        } catch (Exception e) {
            StringBuilder sb = new StringBuilder("詳細差分を生成できませんでした。\n");
            for (Path p : check.mismatchedPaths()) {
                sb.append("・").append(p.getFileName()).append('\n');
            }
            return sb.toString().trim();
        }
    }
}
```

テスト: `summarizeOrFallback` が summarizer 例外時にフォールバックを返すこと。

- [ ] Commit: `feat: 保存競合ゲートの要約フォールバックを追加`

---

### Task 6: MemberAttendanceTabController に組み込み

**Files:**
- Modify: `code_java/src/main/java/jp/co/pm/ai/desktop/MemberAttendanceTabController.java`

- [ ] **Step 1: フィールド追加**

```java
private FingerprintBaseline conflictBaseline;
private final ConflictDiffSummarizer conflictSummarizer = new AttendanceConflictDiffSummarizer();
```

- [ ] **Step 2: 読込成功時に baseline を取る**

`reloadAttendanceDataFromJson` / グリッド適用成功直後で:

```java
private void refreshConflictBaseline() {
    if (shell == null) {
        conflictBaseline = null;
        return;
    }
    try {
        Map<String, String> ui = shell.snapshotUiEnv();
        Path json = AppPaths.attendanceDataJsonPath(ui);
        Path xlsx = AppPaths.attendanceCalendarXlsxPath(ui);
        conflictBaseline = FingerprintBaseline.capture(List.of(json, xlsx));
    } catch (Exception e) {
        conflictBaseline = null;
        // status に軽く出してよい
    }
}
```

セットアップ・復元・再読込の成功パスでも同じメソッドを呼ぶ。

- [ ] **Step 3: `saveEditsAsync` の先頭（4 桁確認の後、temp 作成の前）にゲート**

```java
if (conflictBaseline != null) {
    ConflictCheckResult check = SaveConflictChecker.check(conflictBaseline);
    if (check.kind() == ConflictCheckResult.Kind.IO_ERROR) {
        statusLabel.setText("保存中止: 競合確認に失敗しました — " + check.errorMessage());
        if (onComplete != null) {
            onComplete.accept(false);
        }
        return;
    }
    if (check.kind() == ConflictCheckResult.Kind.CONFLICT) {
        Map<Path, byte[]> disk;
        try {
            disk = SaveConflictGate.readDiskBytes(conflictBaseline);
        } catch (Exception e) {
            statusLabel.setText("保存中止: " + e.getMessage());
            if (onComplete != null) {
                onComplete.accept(false);
            }
            return;
        }
        String summary =
                SaveConflictGate.summarizeOrFallback(
                        conflictSummarizer, conflictBaseline, disk, check);
        Window owner =
                statusLabel.getScene() != null ? statusLabel.getScene().getWindow() : null;
        ConflictSaveChoice choice = ConflictSaveDialog.show(owner, "メンバー勤怠", summary);
        if (choice == ConflictSaveChoice.CANCEL) {
            if (onComplete != null) {
                onComplete.accept(false);
            }
            return;
        }
        if (choice == ConflictSaveChoice.RELOAD) {
            reloadAttendanceDataFromJson();
            if (onComplete != null) {
                onComplete.accept(false);
            }
            return;
        }
        // OVERWRITE → 続行
    }
}
```

- [ ] **Step 4: 保存成功コールバック末尾で `refreshConflictBaseline()`**

- [ ] **Step 5: コンパイル**

Run: `cd code_java && .\mvnw.cmd -q compile`

- [ ] **Step 6: Commit**

```bash
git add code_java/src/main/java/jp/co/pm/ai/desktop/MemberAttendanceTabController.java
git commit -m "feat: メンバー勤怠の保存前外部変更検知を追加"
```

---

### Task 7: CompanyCalendarTabController に組み込み

**Files:**
- Modify: `code_java/src/main/java/jp/co/pm/ai/desktop/CompanyCalendarTabController.java`

Task 6 と同様:

- 同じ JSON + xlsx の `FingerprintBaseline`
- 同じ `AttendanceConflictDiffSummarizer`
- `saveEditsAsync` 相当の入口でゲート
- ダイアログタイトルは「会社カレンダー」
- 再読込は既存の `reloadAttendanceDataFromJson`（または同等）

- [ ] コンパイル確認後 Commit: `feat: 会社カレンダーの保存前外部変更検知を追加`

---

### Task 8: MachineCalendarTabController に組み込み

**Files:**
- Modify: `code_java/src/main/java/jp/co/pm/ai/desktop/MachineCalendarTabController.java`
- Create: `code_java/src/main/java/jp/co/pm/ai/desktop/io/conflict/MachineCalendarConflictDiffSummarizer.java`
- Create: `code_java/src/test/java/jp/co/pm/ai/desktop/io/conflict/MachineCalendarConflictDiffSummarizerTest.java`

- baseline パス: `AppPaths.machineCalendarDataJsonPath(ui)` + `AppPaths.attendanceCalendarXlsxPath(ui)`
- Summarizer: 機械 ID／名称と稼働日区分の差分を業務語で（実 JSON キーに合わせる）。xlsx は「関連 Excel が変更」行。
- ゲート組み込みは Task 6 と同型。タイトル「機械カレンダー」

- [ ] テスト + コンパイル後 Commit: `feat: 機械カレンダーの保存前外部変更検知を追加`

---

### Task 9: 勤怠系の手動確認チェックリスト

- [ ] メンバー勤怠を開き、外部で `attendance-data.json` を編集して保存 → ダイアログ表示、要約に変化が見える
- [ ] 「キャンセル」→ 未保存編集が残る
- [ ] 「再読込して更新」→ UI がディスクに合い、dirty が消える
- [ ] 「強制上書き」→ 画面内容で保存され、その後の再保存では競合しない
- [ ] 関連 xlsx だけ外部変更した場合も競合する

---

## Spec coverage（本計画）

| 仕様項目 | Task |
|----------|------|
| SHA-256 検知 | 1–2 |
| セット指紋 | 2, 6–8 |
| ダイアログ案 B | 3 |
| 業務語要約（勤怠） | 4, 8 |
| 常時ロックしない | 全体（ロック API を追加しない） |
| 会社／メンバー／機械 | 6–8 |
| 他画面 | 後続 2 計画 |

## 次の計画へ

本計画完了後:

1. `2026-09-10-save-conflict-detection-dispatch-screens.md`
2. `2026-09-10-save-conflict-detection-request-and-settings.md`
