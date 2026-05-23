package jp.co.pm.ai.desktop.config;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;

import jp.co.pm.ai.desktop.benchmark.GeminiGenerateContentRestClient;

/**
 * 配台（{@code planning_core}）の Gemini 再試行で使うモデル列の既定。
 *
 * <p>Python 側の {@code GEMINI_MODEL_IDS_BY_QUALITY} と同順を保つこと。
 */
public final class GeminiDispatchModelTryOrderDefaults {

    public static final String ENV_GEMINI_MODEL = "GEMINI_MODEL";
    public static final String ENV_GEMINI_MODEL_TRY_ORDER = "GEMINI_MODEL_TRY_ORDER";

    private GeminiDispatchModelTryOrderDefaults() {}

    /** コード既定の最優先（Flash-Lite 無料枠の自動更新より先に試す）。 */
    public static final String PLANNING_CORE_TOP_PRIORITY_MODEL = "gemini-3.5-flash";

    /** {@code planning_core/_core.py} の {@code GEMINI_MODEL_IDS_BY_QUALITY} と同一。 */
    public static final List<String> PLANNING_CORE_FALLBACK_TRY_ORDER =
            List.of(
                    PLANNING_CORE_TOP_PRIORITY_MODEL,
                    "gemini-3.1-flash-lite",
                    "gemini-3.1-flash-lite-preview",
                    "gemini-2.5-flash-lite",
                    "gemini-2.0-flash-lite");

    /**
     * Flash-Lite 無料枠候補（models.list 等）の前に {@link #PLANNING_CORE_TOP_PRIORITY_MODEL} を置き、
     * 続けてコード既定 Flash-Lite 列を重複除去してマージする。
     */
    public static List<String> withPlanningCorePriorityFirst(List<String> flashLiteCandidates) {
        if (flashLiteCandidates == null || flashLiteCandidates.isEmpty()) {
            return List.copyOf(PLANNING_CORE_FALLBACK_TRY_ORDER);
        }
        LinkedHashSet<String> seen = new LinkedHashSet<>();
        List<String> merged = new ArrayList<>();
        merged.add(PLANNING_CORE_TOP_PRIORITY_MODEL);
        seen.add(PLANNING_CORE_TOP_PRIORITY_MODEL);
        for (String raw : flashLiteCandidates) {
            String norm = GeminiGenerateContentRestClient.normalizeModelId(raw);
            if (!norm.isEmpty() && seen.add(norm)) {
                merged.add(norm);
            }
        }
        for (String id : PLANNING_CORE_FALLBACK_TRY_ORDER) {
            if (seen.add(id)) {
                merged.add(id);
            }
        }
        return List.copyOf(merged);
    }

    /**
     * 環境変数から planning_core と同じ優先で試行モデル列を解決する。
     *
     * <ol>
     *   <li>{@code GEMINI_MODEL} が非空 → その1件
     *   <li>{@code GEMINI_MODEL_TRY_ORDER} が非空 → カンマ区切り（左から）
     *   <li>コード既定 {@link #PLANNING_CORE_FALLBACK_TRY_ORDER}
     * </ol>
     */
    public static List<String> resolveEffectiveModelTryOrder(Map<String, String> env) {
        if (env != null) {
            String pinned = firstNonBlank(env.get(ENV_GEMINI_MODEL));
            if (pinned != null) {
                return List.of(GeminiGenerateContentRestClient.normalizeModelId(pinned));
            }
            List<String> fromOrder = parseTryOrderCsv(env.get(ENV_GEMINI_MODEL_TRY_ORDER));
            if (!fromOrder.isEmpty()) {
                return fromOrder;
            }
        }
        return List.copyOf(PLANNING_CORE_FALLBACK_TRY_ORDER);
    }

    /** カンマ区切り試行順を正規化して重複除去（空なら空リスト）。 */
    public static List<String> parseTryOrderCsv(String raw) {
        if (raw == null || raw.isBlank()) {
            return List.of();
        }
        LinkedHashSet<String> seen = new LinkedHashSet<>();
        List<String> out = new ArrayList<>();
        for (String p : raw.split(",")) {
            if (p == null) {
                continue;
            }
            String t = p.strip();
            if (t.isEmpty()) {
                continue;
            }
            String norm = GeminiGenerateContentRestClient.normalizeModelId(t);
            if (seen.add(norm)) {
                out.add(norm);
            }
        }
        return out;
    }

    private static String firstNonBlank(String value) {
        if (value == null) {
            return null;
        }
        String s = value.strip();
        return s.isEmpty() ? null : s;
    }
}
