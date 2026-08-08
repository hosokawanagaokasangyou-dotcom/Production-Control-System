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
