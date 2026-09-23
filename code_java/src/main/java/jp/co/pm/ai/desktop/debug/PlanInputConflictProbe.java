package jp.co.pm.ai.desktop.debug;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.LinkedHashMap;
import java.util.Map;

import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * 配台計画の保存競合が、どの書き込みのあとで出たかを残す。計測専用。
 */
public final class PlanInputConflictProbe {

    public static final String SESSION = "581499";

    private static final ObjectMapper JSON = new ObjectMapper();
    private static volatile String lastReason = "none";
    private static volatile long lastReasonAt;

    private PlanInputConflictProbe() {}

    public static String reasonText() {
        long agoSec = lastReasonAt == 0 ? -1 : (System.currentTimeMillis() - lastReasonAt) / 1000L;
        String when = agoSec < 0 ? "" : "（" + agoSec + "秒前）";
        String what =
                switch (lastReason) {
                    case "stage1-reload" -> "段階1の完了後に、この画面が計画ファイルを読み込みました";
                    case "stage2-reload" -> "段階2のあと、この画面が計画ファイルを読み込み直しました";
                    case "stage2-reload-skipped-dirty" ->
                            "段階2のあと、未保存の編集があったため計画ファイルを読み直していません";
                    case "ec-side-reloaded" -> "EC面区分を反映したあと、計画ファイルを読み込み直しました";
                    case "dev-exclude-write" -> "開発用の操作が、計画ファイルの配台不要を一括で書きました";
                    case "stage2-unknown-combo" -> "マスタ未登録の工程を配台不要にしたあと、計画ファイルを読み込み直しました";
                    case "user-load" -> "この画面で計画ファイルを読み込みました";
                    case "user-save" -> "この画面で計画ファイルを保存しました";
                    default -> "この画面が計画ファイルを読み込んだ記録があります";
                };
        return what + when;
    }

    public static void note(String reason) {
        lastReason = reason == null ? "none" : reason;
        lastReasonAt = System.currentTimeMillis();
        event("B", "PlanInputConflictProbe.note", "writer", Map.of("reason", lastReason));
    }

    public static void event(
            String hypothesisId, String location, String message, Map<String, ?> data) {
        try {
            Map<String, Object> payload = new LinkedHashMap<>();
            if (data != null) {
                payload.putAll(data);
            }
            payload.put("lastReason", lastReason);
            payload.put("lastReasonAgoMs", lastReasonAt == 0 ? -1 : System.currentTimeMillis() - lastReasonAt);
            Map<String, Object> line = new LinkedHashMap<>();
            line.put("sessionId", SESSION);
            line.put("hypothesisId", hypothesisId);
            line.put("location", location);
            line.put("message", message);
            line.put("data", payload);
            line.put("timestamp", System.currentTimeMillis());
            String json = JSON.writeValueAsString(line);
            AgentDebugLog.appendNdjsonLine(Map.of(), SESSION, json);
            appendWorkspace(json);
        } catch (Throwable ignored) {
            // 計測失敗で保存は止めない
        }
    }

    private static void appendWorkspace(String json) {
        Path repo = findRepoRoot();
        if (repo == null) {
            return;
        }
        try {
            Files.writeString(
                    repo.resolve("debug-581499.log"),
                    json + "\n",
                    StandardCharsets.UTF_8,
                    StandardOpenOption.CREATE,
                    StandardOpenOption.APPEND);
        } catch (Exception ignored) {
            // 計測失敗で保存は止めない
        }
    }

    private static Path findRepoRoot() {
        Path dir = Path.of(System.getProperty("user.dir", ".")).toAbsolutePath().normalize();
        for (int i = 0; i < 8 && dir != null; i++) {
            if (Files.isRegularFile(dir.resolve("version.txt"))
                    && Files.isDirectory(dir.resolve("code_java"))) {
                return dir;
            }
            dir = dir.getParent();
        }
        return null;
    }
}
