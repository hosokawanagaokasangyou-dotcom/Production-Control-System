# PmAiFxApp 二重起動抑制 Implementation Plan

> **For implementers:** Execute this plan task-by-task. Use the `test-driven-development` skill for each task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** `PmAiFxApp` を同一マシンで 1 インスタンスに制限し、2つ目起動時は既存主窓を前面化してダイアログなしで終了する。

**Architecture:** `127.0.0.1` 固定ポートのローカルソケットで Primary / Secondary を判定する。`SingleInstanceGuard` が listen・ACTIVATE プロトコル・Stage 前面化コールバックを担当し、`PmAiFxApp.main` の早期でガードする（GPU プローブ前）。無効化は `-Dpm.ai.singleInstance=false`。

**Tech Stack:** Java 17+、`java.net.ServerSocket` / `Socket`、JavaFX `Stage` / `Platform`、JUnit 5、Maven（`code_java/mvnw.cmd`）

**仕様正本:** `docs/specs/2026-08-08-single-instance-guard-design.md`

---

## ファイル構成

| ファイル | 責任 |
|----------|------|
| `code_java/src/main/java/jp/co/pm/ai/desktop/runtime/SingleInstanceGuard.java` | ポート解決・Secondary 試行・Primary listen・終了時 close・前面化コールバック登録 |
| `code_java/src/test/java/jp/co/pm/ai/desktop/runtime/SingleInstanceGuardTest.java` | ヘッドレス単体テスト |
| `code_java/src/main/java/jp/co/pm/ai/desktop/PmAiFxApp.java` | `main` 早期ガード、`start` で Stage 登録、終了時解放 |

---

### Task 1: SingleInstanceGuard — ACTIVATE プロトコルと Primary listen

**Files:**
- Create: `code_java/src/test/java/jp/co/pm/ai/desktop/runtime/SingleInstanceGuardTest.java`
- Create: `code_java/src/main/java/jp/co/pm/ai/desktop/runtime/SingleInstanceGuard.java`

- [ ] **Step 1: Write the failing test**

`SingleInstanceGuardTest.java` を作成する（クラス未存在でコンパイル失敗／テスト失敗になる）。

```java
package jp.co.pm.ai.desktop.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicInteger;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

class SingleInstanceGuardTest {

    private SingleInstanceGuard guard;
    private int port;

    @BeforeEach
    void setUp() throws Exception {
        port = SingleInstanceGuard.findFreePort();
        System.setProperty(SingleInstanceGuard.PROP_ENABLED, "true");
        System.setProperty(SingleInstanceGuard.PROP_PORT, Integer.toString(port));
        guard = new SingleInstanceGuard();
    }

    @AfterEach
    void tearDown() {
        if (guard != null) {
            guard.close();
        }
        System.clearProperty(SingleInstanceGuard.PROP_ENABLED);
        System.clearProperty(SingleInstanceGuard.PROP_PORT);
    }

    @Test
    void primaryAcceptsActivateAndInvokesCallbackOnce() throws Exception {
        AtomicInteger activations = new AtomicInteger();
        CountDownLatch latch = new CountDownLatch(1);
        guard.setOnActivateRequest(
                () -> {
                    activations.incrementAndGet();
                    latch.countDown();
                });

        assertEquals(SingleInstanceGuard.Role.PRIMARY, guard.tryAcquire());

        assertTrue(SingleInstanceGuard.sendActivate(port, 500));
        assertTrue(latch.await(2, TimeUnit.SECONDS));
        assertEquals(1, activations.get());
    }

    @Test
    void secondAcquireBecomesSecondaryWhenPrimaryListening() throws Exception {
        assertEquals(SingleInstanceGuard.Role.PRIMARY, guard.tryAcquire());

        SingleInstanceGuard second = new SingleInstanceGuard();
        try {
            assertEquals(SingleInstanceGuard.Role.SECONDARY, second.tryAcquire());
        } finally {
            second.close();
        }
    }

    @Test
    void disabledPropertySkipsGuard() throws Exception {
        System.setProperty(SingleInstanceGuard.PROP_ENABLED, "false");
        assertEquals(SingleInstanceGuard.Role.DISABLED, guard.tryAcquire());
        assertFalse(SingleInstanceGuard.sendActivate(port, 200));
    }
}
```

- [ ] **Step 2: Run test to verify it fails**

```powershell
cd code_java
.\mvnw.cmd -q test -Dtest=SingleInstanceGuardTest
```

Expected: コンパイル失敗（`SingleInstanceGuard` が無い）またはテスト失敗。

- [ ] **Step 3: Write minimal implementation**

`SingleInstanceGuard.java`:

```java
package jp.co.pm.ai.desktop.runtime;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.net.InetAddress;
import java.net.ServerSocket;
import java.net.Socket;
import java.net.SocketTimeoutException;
import java.nio.charset.StandardCharsets;
import java.util.concurrent.atomic.AtomicReference;

/**
 * PmAiFxApp 用の単一インスタンス制御（127.0.0.1 ソケット）。
 *
 * <p>無効化: {@code -Dpm.ai.singleInstance=false}／ポート: {@code -Dpm.ai.singleInstance.port}
 */
public final class SingleInstanceGuard implements AutoCloseable {

    public static final String PROP_ENABLED = "pm.ai.singleInstance";
    public static final String PROP_PORT = "pm.ai.singleInstance.port";
    public static final int DEFAULT_PORT = 47821;
    public static final String ACTIVATE_CMD = "ACTIVATE";
    public static final String OK_RESP = "OK";

    public enum Role {
        PRIMARY,
        SECONDARY,
        DISABLED,
        /** bind 失敗などでガード不能。呼び出し側は通常起動してよい */
        UNAVAILABLE
    }

    private final AtomicReference<Runnable> onActivate = new AtomicReference<>();
    private volatile ServerSocket server;
    private volatile Thread acceptThread;

    public void setOnActivateRequest(Runnable callback) {
        onActivate.set(callback);
    }

    public Role tryAcquire() {
        if (!isEnabled()) {
            return Role.DISABLED;
        }
        int port = resolvePort();
        if (sendActivate(port, 300)) {
            return Role.SECONDARY;
        }
        try {
            ServerSocket ss = new ServerSocket(port, 1, InetAddress.getByName("127.0.0.1"));
            server = ss;
            acceptThread = new Thread(this::acceptLoop, "pm-ai-single-instance");
            acceptThread.setDaemon(true);
            acceptThread.start();
            return Role.PRIMARY;
        } catch (IOException e) {
            return Role.UNAVAILABLE;
        }
    }

    public static boolean isEnabled() {
        String raw = System.getProperty(PROP_ENABLED);
        if (raw == null || raw.isBlank()) {
            return true;
        }
        return !"false".equalsIgnoreCase(raw.trim())
                && !"0".equals(raw.trim())
                && !"off".equalsIgnoreCase(raw.trim());
    }

    public static int resolvePort() {
        String raw = System.getProperty(PROP_PORT);
        if (raw == null || raw.isBlank()) {
            return DEFAULT_PORT;
        }
        try {
            int p = Integer.parseInt(raw.trim());
            return p > 0 && p <= 65535 ? p : DEFAULT_PORT;
        } catch (NumberFormatException e) {
            return DEFAULT_PORT;
        }
    }

    /** テスト用: OS が割り当てた空きポート。 */
    public static int findFreePort() throws IOException {
        try (ServerSocket ss = new ServerSocket(0, 1, InetAddress.getByName("127.0.0.1"))) {
            return ss.getLocalPort();
        }
    }

    public static boolean sendActivate(int port, int timeoutMs) {
        try (Socket socket = new Socket()) {
            socket.connect(
                    new java.net.InetSocketAddress(InetAddress.getByName("127.0.0.1"), port),
                    timeoutMs);
            socket.setSoTimeout(timeoutMs);
            PrintWriter out =
                    new PrintWriter(
                            new OutputStreamWriter(socket.getOutputStream(), StandardCharsets.UTF_8),
                            true);
            BufferedReader in =
                    new BufferedReader(
                            new InputStreamReader(socket.getInputStream(), StandardCharsets.UTF_8));
            out.println(ACTIVATE_CMD);
            String line = in.readLine();
            return OK_RESP.equals(line);
        } catch (IOException e) {
            return false;
        }
    }

    private void acceptLoop() {
        ServerSocket ss = server;
        if (ss == null) {
            return;
        }
        while (!ss.isClosed()) {
            try (Socket client = ss.accept()) {
                handleClient(client);
            } catch (SocketTimeoutException ignored) {
                /* unused */
            } catch (IOException e) {
                if (ss.isClosed()) {
                    break;
                }
            }
        }
    }

    private void handleClient(Socket client) throws IOException {
        client.setSoTimeout(1000);
        BufferedReader in =
                new BufferedReader(
                        new InputStreamReader(client.getInputStream(), StandardCharsets.UTF_8));
        PrintWriter out =
                new PrintWriter(
                        new OutputStreamWriter(client.getOutputStream(), StandardCharsets.UTF_8),
                        true);
        String line = in.readLine();
        if (ACTIVATE_CMD.equals(line)) {
            out.println(OK_RESP);
            Runnable cb = onActivate.get();
            if (cb != null) {
                cb.run();
            }
        }
    }

    @Override
    public void close() {
        ServerSocket ss = server;
        server = null;
        if (ss != null) {
            try {
                ss.close();
            } catch (IOException ignored) {
                /* ignore */
            }
        }
        Thread t = acceptThread;
        acceptThread = null;
        if (t != null) {
            try {
                t.join(500);
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
            }
        }
    }
}
```

- [ ] **Step 4: Run tests to verify they pass**

```powershell
cd code_java
.\mvnw.cmd -q test -Dtest=SingleInstanceGuardTest
```

Expected: BUILD SUCCESS / tests pass。

- [ ] **Step 5: Commit**（ユーザーがコミット不要と言わない限り）

```powershell
git add code_java/src/main/java/jp/co/pm/ai/desktop/runtime/SingleInstanceGuard.java `
  code_java/src/test/java/jp/co/pm/ai/desktop/runtime/SingleInstanceGuardTest.java
git commit -m "feat: PmAiFxApp用 SingleInstanceGuard を追加"
```

---

### Task 2: PmAiFxApp への配線（早期ガード・Stage 前面化・終了解放）

**Files:**
- Modify: `code_java/src/main/java/jp/co/pm/ai/desktop/PmAiFxApp.java`
- Modify: `code_java/src/test/java/jp/co/pm/ai/desktop/runtime/SingleInstanceGuardTest.java`（任意で配線用の契約テストは不要。手動確認が主）

- [ ] **Step 1: Write a failing contract test that `PmAiFxApp` references the guard**

`code_java/src/test/java/jp/co/pm/ai/desktop/PmAiFxAppSingleInstanceContractTest.java` を追加:

```java
package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;

class PmAiFxAppSingleInstanceContractTest {

    @Test
    void mainSourceWiresSingleInstanceGuardEarly() throws Exception {
        Path src =
                Path.of("src/main/java/jp/co/pm/ai/desktop/PmAiFxApp.java");
        String text = Files.readString(src, StandardCharsets.UTF_8);
        assertTrue(text.contains("SingleInstanceGuard"));
        assertTrue(text.contains("tryAcquire"));
        assertTrue(text.contains("Role.SECONDARY"));
        assertTrue(text.contains("setOnActivateRequest"));
        assertTrue(text.contains("bringPrimaryStageToFront"));
    }
}
```

（テスト作業ディレクトリは Maven surefire 既定で `code_java`。）

- [ ] **Step 2: Run contract test — expect FAIL**

```powershell
cd code_java
.\mvnw.cmd -q test -Dtest=PmAiFxAppSingleInstanceContractTest
```

Expected: FAIL（ソースに未配線）。

- [ ] **Step 3: Wire `PmAiFxApp`**

変更要点:

1. フィールド追加:

```java
private static final SingleInstanceGuard SINGLE_INSTANCE = new SingleInstanceGuard();
private static volatile Stage primaryStageRef;
```

2. `main` 内、headless チェックの直後・`configurePrismAfterProbe()` の前:

```java
SINGLE_INSTANCE.setOnActivateRequest(PmAiFxApp::bringPrimaryStageToFront);
SingleInstanceGuard.Role role = SINGLE_INSTANCE.tryAcquire();
if (role == SingleInstanceGuard.Role.SECONDARY) {
    StartupCrashLog.append("main: secondary instance — activate requested, exit");
    System.exit(0);
    return;
}
if (role == SingleInstanceGuard.Role.UNAVAILABLE) {
    StartupCrashLog.append(
            "main: single-instance guard unavailable (port busy?) — continuing");
}
Runtime.getRuntime().addShutdownHook(new Thread(SINGLE_INSTANCE::close, "pm-ai-single-instance-shutdown"));
```

3. `start(Stage primaryStage)` の先頭付近で:

```java
primaryStageRef = primaryStage;
primaryStage.setOnHidden(e -> SINGLE_INSTANCE.close());
```

（`setOnCloseRequest` だけだと cancel される場合があるため、実際に隠れたタイミングで close。shutdown hook も併用。）

4. 前面化ヘルパ:

```java
private static void bringPrimaryStageToFront() {
    Platform.runLater(
            () -> {
                Stage stage = primaryStageRef;
                if (stage == null) {
                    return;
                }
                if (stage.isIconified()) {
                    stage.setIconified(false);
                }
                stage.toFront();
                stage.requestFocus();
            });
}
```

import 追加:

```java
import jp.co.pm.ai.desktop.runtime.SingleInstanceGuard;
```

- [ ] **Step 4: Run tests**

```powershell
cd code_java
.\mvnw.cmd -q test -Dtest=SingleInstanceGuardTest,PmAiFxAppSingleInstanceContractTest
```

Expected: PASS。

- [ ] **Step 5: 手動確認（実装者）**

1. 通常起動スクリプトで 1 つ目を起動し主窓表示を待つ  
2. 同じ方法で 2 つ目を起動 → 1つ目が前面化し、2つ目プロセスがすぐ終了  
3. `-Dpm.ai.singleInstance=false` 付きで二重起動できること

- [ ] **Step 6: Commit**

```powershell
git add code_java/src/main/java/jp/co/pm/ai/desktop/PmAiFxApp.java `
  code_java/src/test/java/jp/co/pm/ai/desktop/PmAiFxAppSingleInstanceContractTest.java
git commit -m "feat: PmAiFxApp に二重起動抑制を配線"
```

---

### Task 3: 仕様ステータス更新

**Files:**
- Modify: `docs/specs/2026-08-08-single-instance-guard-design.md`

- [ ] **Step 1:** 先頭の `状態:` を `実装完了` に変更する。

- [ ] **Step 2: Commit / push**（リポジトリ方針に従う）

```powershell
git add docs/specs/2026-08-08-single-instance-guard-design.md
git commit -m "docs: 二重起動抑制の設計メモを実装完了に更新"
git push
```

---

## Spec coverage（自己レビュー）

| 仕様要件 | Task |
|----------|------|
| PmAiFxApp のみ | Task 2（他 App は触らない） |
| 前面化＋ダイアログなし exit | Task 1 コールバック + Task 2 `bringPrimaryStageToFront` / `System.exit(0)` |
| `pm.ai.singleInstance=false` | Task 1 `DISABLED` テスト |
| ポート `47821` / 上書き | Task 1 `DEFAULT_PORT` / `PROP_PORT` |
| GPU プローブ前 | Task 2 `main` 配置 |
| bind 失敗で通常起動＋ログ | Task 2 `UNAVAILABLE` + `StartupCrashLog` |
| ServerSocket 解放 | Task 1 `close` + Task 2 shutdown hook / `setOnHidden` |
| 単体テスト | Task 1 |
| 手動確認 | Task 2 Step 5 |
