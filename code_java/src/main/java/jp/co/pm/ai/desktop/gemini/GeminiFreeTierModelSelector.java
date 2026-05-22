package jp.co.pm.ai.desktop.gemini;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import jp.co.pm.ai.desktop.benchmark.GeminiGenerateContentRestClient;
import jp.co.pm.ai.desktop.benchmark.GeminiModelsListRestClient.ListedModel;

/**
 * {@code models.list} 結果から、無料枠運用向けの Flash-Lite 系（{@code generateContent} 対応）を抽出・並べ替えする。
 */
public final class GeminiFreeTierModelSelector {

    private static final Pattern VERSION_PREFIX =
            Pattern.compile("^gemini-(\\d+)(?:\\.(\\d+))?");

    private GeminiFreeTierModelSelector() {}

    /**
     * Flash-Lite かつ {@code generateContent} をサポートするモデル ID を、新しい世代→旧、同一世代は GA→preview の順で返す。
     */
    public static List<String> selectFlashLiteGenerateContentModels(List<ListedModel> listed) {
        if (listed == null || listed.isEmpty()) {
            return List.of();
        }
        LinkedHashSet<String> seen = new LinkedHashSet<>();
        List<String> candidates = new ArrayList<>();
        for (ListedModel m : listed) {
            if (m == null) {
                continue;
            }
            if (!supportsGenerateContent(m.supportedGenerationMethods())) {
                continue;
            }
            String id = GeminiGenerateContentRestClient.normalizeModelId(m.modelId());
            if (id.isEmpty() || !GeminiGenerateContentRestClient.isAllowedModelId(id)) {
                continue;
            }
            String lower = id.toLowerCase(Locale.ROOT);
            if (!lower.contains("flash-lite") && !lower.contains("flashlite")) {
                continue;
            }
            if (isExcludedSuffix(lower)) {
                continue;
            }
            if (seen.add(id)) {
                candidates.add(id);
            }
        }
        candidates.sort(GeminiFreeTierModelSelector::compareModelIds);
        return List.copyOf(candidates);
    }

    private static boolean supportsGenerateContent(List<String> methods) {
        if (methods == null) {
            return false;
        }
        for (String m : methods) {
            if ("generateContent".equals(m)) {
                return true;
            }
        }
        return false;
    }

    private static boolean isExcludedSuffix(String lowerId) {
        return lowerId.contains("embedding")
                || lowerId.contains("embed")
                || lowerId.contains("aqa")
                || lowerId.contains("imagen")
                || lowerId.contains("veo")
                || lowerId.contains("live")
                || lowerId.contains("tts")
                || lowerId.contains("robotics");
    }

    /** 新しい gemini メジャー／マイナーを先に。同一系統は preview より非 preview を先。 */
    static int compareModelIds(String a, String b) {
        VersionKey ka = versionKey(a);
        VersionKey kb = versionKey(b);
        int c = Integer.compare(kb.major, ka.major);
        if (c != 0) {
            return c;
        }
        c = Integer.compare(kb.minor, ka.minor);
        if (c != 0) {
            return c;
        }
        c = Boolean.compare(ka.preview, kb.preview);
        if (c != 0) {
            return c;
        }
        return a.compareTo(b);
    }

    private static VersionKey versionKey(String modelId) {
        String lower = modelId.toLowerCase(Locale.ROOT);
        boolean preview = lower.contains("preview") || lower.contains("-exp");
        Matcher m = VERSION_PREFIX.matcher(lower);
        int major = 0;
        int minor = 0;
        if (m.find()) {
            major = parseIntSafe(m.group(1));
            if (m.group(2) != null) {
                minor = parseIntSafe(m.group(2));
            }
        }
        return new VersionKey(major, minor, preview);
    }

    private static int parseIntSafe(String s) {
        try {
            return Integer.parseInt(s);
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    private record VersionKey(int major, int minor, boolean preview) {}
}
