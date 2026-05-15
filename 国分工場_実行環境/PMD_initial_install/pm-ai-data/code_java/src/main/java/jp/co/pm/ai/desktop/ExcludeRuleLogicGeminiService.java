package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

import jp.co.pm.ai.desktop.benchmark.GeminiGenerateContentRestClient;
import jp.co.pm.ai.desktop.benchmark.GeminiGenerateContentRestClient.FullBodyResult;
import jp.co.pm.ai.desktop.config.GeminiDispatchModelTryOrderDefaults;

/**
 * 「配台不要ロジック」自然言語を Gemini で {@code ロジック式} JSON に変換する（{@code planning_core._core} の評価器と整合）。
 */
public final class ExcludeRuleLogicGeminiService {

    private static final ObjectMapper COMPACT_JSON =
            new ObjectMapper().disable(SerializationFeature.INDENT_OUTPUT);

    private static final Pattern JSON_FENCE =
            Pattern.compile("```(?:json)?\\s*(\\{.*\\})\\s*```", Pattern.CASE_INSENSITIVE | Pattern.DOTALL);

    private ExcludeRuleLogicGeminiService() {}

    /**
     * {@code planning_core._core.EXCLUDE_RULE_ALLOWED_COLUMNS} と同一（列名は評価器が拒否しないもののみ）。
     */
    private static List<String> canonicalAllowedList() {
        List<String> u = new ArrayList<>();
        u.add("(原反)ロール単位長さ");
        u.add("データ抽出日");
        u.add("データ抽出時間");
        u.add("使用原反");
        u.add("加工内容");
        u.add("加工完了区分");
        u.add("加工工程の決定プロセスの因子");
        u.add("加工速度");
        u.add("加工速度_上書き");
        u.add("原反幅");
        u.add("原反投入日");
        u.add("原反投入日_上書き");
        u.add("受注数");
        u.add("回答納期");
        u.add("在庫場所");
        u.add("実加工数");
        u.add("実出来高");
        u.add("指定納期");
        u.add("担当OP_指定");
        u.add("換算数量");
        u.add("抽出時間");
        u.add("未加工");
        u.add("機械名");
        u.add("工程名");
        u.add("依頼NO");
        u.add("製品名");
        u.add("製品厚み");
        u.add("製品幅");
        u.add("製品長");
        u.add("特別指定_備考");
        u.add("ロール単位長さ");
        u.add("配台使用残数量");
        u.sort(Comparator.naturalOrder());
        return List.copyOf(u);
    }

    private static String allowedColumnsCsv() {
        return String.join(", ", canonicalAllowedList());
    }

    private static String schemaInstructions() {
        return "【スキーマ】version は必ず 1。\n"
                + "1) 常に配台試行から外す意味のとき:\n"
                + "{\"version\":1,\"mode\":\"always_exclude\"}\n\n"
                + "2) 列の条件で配台試行から外すとき:\n"
                + "{\"version\":1,\"mode\":\"conditions\",\"require_all\": true または false,\"conditions\":[ ... ]}\n\n"
                + "conditions の各要素:\n"
                + "- {\"column\":\"列名\",\"op\":\"empty\"} … セルは空\n"
                + "- {\"column\":\"列名\",\"op\":\"not_empty\"}\n"
                + "- {\"column\":\"列名\",\"op\":\"eq\",\"value\":\"文字列\"} / ne / contains / not_contains / regex（正規表現）\n"
                + "- {\"column\":\"列名\",\"op\":\"gt\"|\"gte\"|\"lt\"|\"lte\",\"value\":数値} … 数値比較（列は数として解釈）\n\n"
                + "【使用可能な列名のみ】（これ以外は使えない）:\n"
                + allowedColumnsCsv()
                + "\n";
    }

    private static String buildPrompt(String naturalLanguage, String processName, String machineName) {
        StringBuilder ctx = new StringBuilder();
        if (processName != null && !processName.isBlank()) {
            ctx.append("（当該ルール行の工程名: ").append(processName.strip()).append("）\n");
        }
        if (machineName != null && !machineName.isBlank()) {
            ctx.append("（当該ルール行の機械名: ").append(machineName.strip()).append("）\n");
        }
        return "あなたは工場の配台システム用です。次の「配台試行の説明」を、タスク1行を判定する機械可読ルール（JSON）に変換してください。\n\n"
                + "【出力】先頭は { で終わりは } の JSON オブジェクト1つのみ（説明文・マークダウン・コードフェンスは禁止）。\n\n"
                + schemaInstructions()
                + "\n"
                + ctx
                + "【説明文】\n"
                + naturalLanguage.strip()
                + "\n";
    }

    static JsonNode parseModelJsonObject(String raw) throws IOException {
        String s = raw != null ? raw.strip() : "";
        if (s.isEmpty()) {
            throw new IOException("モデル応答が空です。");
        }
        Matcher fence = JSON_FENCE.matcher(s);
        if (fence.find()) {
            s = fence.group(1).strip();
        }
        try {
            JsonNode n = COMPACT_JSON.readTree(s);
            if (n.isObject()) {
                return n;
            }
        } catch (IOException ignored) {
            // fall through
        }
        int i = s.indexOf('{');
        int j = s.lastIndexOf('}');
        if (i >= 0 && j > i) {
            JsonNode n = COMPACT_JSON.readTree(s.substring(i, j + 1));
            if (n.isObject()) {
                return n;
            }
        }
        throw new IOException("モデル応答から JSON オブジェクトを解釈できません。");
    }

    static JsonNode validateRuleJson(JsonNode o) throws IOException {
        if (o == null || !o.isObject()) {
            throw new IOException("ルールは JSON オブジェクトである必要があります。");
        }
        if (o.path("version").asInt(0) != 1) {
            throw new IOException("version は 1 である必要があります。");
        }
        String mode = o.path("mode").asText("").strip().toLowerCase(Locale.ROOT);
        if (!"always_exclude".equals(mode) && !"conditions".equals(mode)) {
            throw new IOException("mode は always_exclude または conditions である必要があります。");
        }
        return o;
    }

    /**
     * Gemini で自然言語をルール JSON テキスト（1行・インデントなし）に変換する。
     *
     * @param naturalLanguage 「配台不要ロジック」列の本文
     * @param processName 行の工程名（任意・プロンプト補助）
     * @param machineName 行の機械名（任意）
     * @param apiKey 復号済み API キー
     * @param timeout 1 モデルあたりのタイムアウト
     * @return コンパクト JSON 文字列
     */
    public static String compileToCompactJson(
            String naturalLanguage,
            String processName,
            String machineName,
            String apiKey,
            Duration timeout)
            throws IOException, InterruptedException {
        String blob = naturalLanguage != null ? naturalLanguage.strip() : "";
        if (blob.isEmpty()) {
            throw new IOException("配台不要ロジックが空です。");
        }
        String prompt = buildPrompt(blob, processName, machineName);
        IOException lastIo = null;
        String lastErr = "";
        for (String modelId : GeminiDispatchModelTryOrderDefaults.PLANNING_CORE_FALLBACK_TRY_ORDER) {
            FullBodyResult res =
                    GeminiGenerateContentRestClient.generateContentFullBody(
                            apiKey, modelId, prompt, 4096, timeout);
            if (!res.errorSummary().isEmpty()) {
                lastErr = res.errorSummary();
                int st = res.httpStatus();
                if (st == 404 || st == 429) {
                    continue;
                }
                throw new IOException(lastErr);
            }
            Optional<String> textOpt =
                    GeminiGenerateContentRestClient.extractFirstCandidateText(res.body());
            if (textOpt.isEmpty()) {
                lastErr = "候補テキストが空です（model=" + modelId + "）";
                continue;
            }
            try {
                JsonNode parsed = validateRuleJson(parseModelJsonObject(textOpt.get()));
                return COMPACT_JSON.writeValueAsString(parsed);
            } catch (IOException ex) {
                lastIo = ex;
                lastErr = ex.getMessage();
            }
        }
        if (lastIo != null) {
            throw new IOException(lastErr != null ? lastErr : lastIo.getMessage(), lastIo);
        }
        throw new IOException(lastErr != null ? lastErr : "Gemini 呼び出しに失敗しました。");
    }
}
