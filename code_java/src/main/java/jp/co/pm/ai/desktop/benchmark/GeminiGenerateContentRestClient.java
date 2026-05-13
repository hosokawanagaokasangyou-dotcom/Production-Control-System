package jp.co.pm.ai.desktop.benchmark;

import java.io.IOException;
import java.net.URI;
import java.net.URLEncoder;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.time.Duration;
import java.util.Objects;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * Google Gemini {@code generateContent}（REST v1beta）を呼び出し、往復時間を計測する。
 *
 * <p>認証はクエリ {@code key=}（API キー）方式。
 */
public final class GeminiGenerateContentRestClient {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    private GeminiGenerateContentRestClient() {}

    /**
     * @param wallTimeNanos {@link System#nanoTime()} ベースの往復時間
     * @param bodyPreview 応答本文の先頭（ログ用・長さは {@link #preview} で制限）
     */
    public record CallResult(int httpStatus, long wallTimeNanos, String bodyPreview, String errorSummary) {

        public double wallTimeMs() {
            return wallTimeNanos / 1_000_000.0;
        }
    }

    /**
     * {@code models/…} 形式や前後空白を正規化する。
     */
    public static String normalizeModelId(String raw) {
        String s = raw != null ? raw.strip() : "";
        if (s.startsWith("models/")) {
            s = s.substring("models/".length()).strip();
        }
        return s;
    }

    /**
     * Gemini に 1 リクエスト送る。
     *
     * @param requestTimeout 接続＋応答待ちの上限（{@link HttpRequest.Builder#timeout} に使用）
     */
    public static CallResult generateContent(
            String apiKey,
            String modelId,
            String userPrompt,
            int maxOutputTokens,
            Duration requestTimeout)
            throws IOException, InterruptedException {
        Objects.requireNonNull(apiKey, "apiKey");
        String model = normalizeModelId(modelId);
        if (model.isEmpty()) {
            throw new IllegalArgumentException("モデル ID が空です。");
        }
        if (!model.matches("^[a-zA-Z0-9._\\-]+$")) {
            throw new IllegalArgumentException("モデル ID は英数字・ドット・ハイフン・アンダースコアのみ使用してください。");
        }
        String key = apiKey.strip();
        if (key.isEmpty()) {
            throw new IllegalArgumentException("API キーが空です。");
        }
        String encKey = URLEncoder.encode(key, StandardCharsets.UTF_8);
        // モデル ID はパスセグメントとして安全な文字のみ想定（ドットはエンコードしない）
        URI uri =
                URI.create(
                        "https://generativelanguage.googleapis.com/v1beta/models/"
                                + model
                                + ":generateContent?key="
                                + encKey);

        String jsonBody = buildRequestJson(userPrompt, maxOutputTokens);

        HttpClient client =
                HttpClient.newBuilder().connectTimeout(requestTimeout).build();
        HttpRequest req =
                HttpRequest.newBuilder(uri)
                        .timeout(requestTimeout)
                        .header("Content-Type", "application/json; charset=utf-8")
                        .POST(HttpRequest.BodyPublishers.ofString(jsonBody, StandardCharsets.UTF_8))
                        .build();

        long t0 = System.nanoTime();
        HttpResponse<String> res = client.send(req, HttpResponse.BodyHandlers.ofString(StandardCharsets.UTF_8));
        long dt = System.nanoTime() - t0;

        int code = res.statusCode();
        String body = res.body() != null ? res.body() : "";
        String preview = preview(body, 480);
        String err = "";
        if (code < 200 || code >= 300) {
            err = summarizeHttpError(code, body);
        }
        return new CallResult(code, dt, preview, err);
    }

    private static String buildRequestJson(String userPrompt, int maxOutputTokens) throws IOException {
        ObjectNode root = JsonNodeFactory.instance.objectNode();
        ObjectNode part = JsonNodeFactory.instance.objectNode();
        part.put("text", userPrompt != null ? userPrompt : "");
        ArrayNode parts = JsonNodeFactory.instance.arrayNode();
        parts.add(part);
        ObjectNode content = JsonNodeFactory.instance.objectNode();
        content.set("parts", parts);
        ArrayNode contents = JsonNodeFactory.instance.arrayNode();
        contents.add(content);
        root.set("contents", contents);
        ObjectNode gen = JsonNodeFactory.instance.objectNode();
        gen.put("maxOutputTokens", Math.max(1, Math.min(maxOutputTokens, 8192)));
        gen.put("temperature", 0.0);
        root.set("generationConfig", gen);
        return MAPPER.writeValueAsString(root);
    }

    private static String preview(String body, int maxChars) {
        if (body == null || body.isEmpty()) {
            return "";
        }
        String oneLine = body.replace('\r', ' ').replace('\n', ' ').strip();
        if (oneLine.length() <= maxChars) {
            return oneLine;
        }
        return oneLine.substring(0, maxChars) + "…";
    }

    private static String summarizeHttpError(int code, String body) {
        String p = preview(body, 240);
        if (!p.isEmpty()) {
            return "HTTP " + code + ": " + p;
        }
        return "HTTP " + code;
    }
}
