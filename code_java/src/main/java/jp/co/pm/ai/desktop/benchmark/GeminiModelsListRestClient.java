package jp.co.pm.ai.desktop.benchmark;

import java.io.IOException;
import java.net.URI;
import java.net.URLEncoder;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * Google Gemini {@code models.list}（REST v1beta）で利用可能モデルを列挙する。
 */
public final class GeminiModelsListRestClient {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    private static final String LIST_BASE =
            "https://generativelanguage.googleapis.com/v1beta/models";

    private GeminiModelsListRestClient() {}

    /** {@code models.list} の1件分（必要フィールドのみ）。 */
    public record ListedModel(String name, List<String> supportedGenerationMethods) {

        /** {@code models/gemini-2.5-flash-lite} → {@code gemini-2.5-flash-lite} */
        public String modelId() {
            return GeminiGenerateContentRestClient.normalizeModelId(name);
        }
    }

    /**
     * 全ページを辿ってモデル一覧を返す。
     *
     * @throws IOException HTTP 失敗・JSON 不正
     */
    public static List<ListedModel> listAllModels(String apiKey, Duration requestTimeout)
            throws IOException, InterruptedException {
        Objects.requireNonNull(apiKey, "apiKey");
        String key = apiKey.strip();
        if (key.isEmpty()) {
            throw new IllegalArgumentException("API キーが空です。");
        }
        Duration timeout = requestTimeout != null ? requestTimeout : Duration.ofSeconds(60);
        String encKey = URLEncoder.encode(key, StandardCharsets.UTF_8);
        HttpClient client =
                HttpClient.newBuilder().connectTimeout(timeout).build();

        List<ListedModel> out = new ArrayList<>();
        String pageToken = null;
        do {
            StringBuilder url = new StringBuilder(LIST_BASE).append("?key=").append(encKey).append("&pageSize=100");
            if (pageToken != null && !pageToken.isBlank()) {
                url.append("&pageToken=").append(URLEncoder.encode(pageToken, StandardCharsets.UTF_8));
            }
            HttpRequest req =
                    HttpRequest.newBuilder(URI.create(url.toString()))
                            .timeout(timeout)
                            .GET()
                            .build();
            HttpResponse<String> res =
                    client.send(req, HttpResponse.BodyHandlers.ofString(StandardCharsets.UTF_8));
            int code = res.statusCode();
            String body = res.body() != null ? res.body() : "";
            if (code < 200 || code >= 300) {
                throw new IOException(
                        "models.list HTTP " + code + (body.isBlank() ? "" : ": " + preview(body, 240)));
            }
            JsonNode root = MAPPER.readTree(body);
            JsonNode models = root.path("models");
            if (models.isArray()) {
                for (JsonNode m : models) {
                    String name = m.path("name").asText("");
                    if (name.isBlank()) {
                        continue;
                    }
                    List<String> methods = new ArrayList<>();
                    JsonNode sm = m.path("supportedGenerationMethods");
                    if (sm.isArray()) {
                        for (JsonNode item : sm) {
                            methods.add(item.asText(""));
                        }
                    }
                    out.add(new ListedModel(name, List.copyOf(methods)));
                }
            }
            JsonNode next = root.path("nextPageToken");
            pageToken = next.isMissingNode() || next.isNull() ? null : next.asText("");
        } while (pageToken != null && !pageToken.isBlank());
        return List.copyOf(out);
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
}
