package jp.co.pm.ai.desktop.config;

import java.util.List;

/**
 * 配台（{@code planning_core}）の Gemini 再試行で使うモデル列の既定。
 *
 * <p>Python 側の {@code GEMINI_MODEL_IDS_BY_QUALITY} と同順を保つこと。
 */
public final class GeminiDispatchModelTryOrderDefaults {

    private GeminiDispatchModelTryOrderDefaults() {}

    /** {@code planning_core/_core.py} の {@code GEMINI_MODEL_IDS_BY_QUALITY} と同一。 */
    public static final List<String> PLANNING_CORE_FALLBACK_TRY_ORDER =
            List.of(
                    "gemini-3.1-flash-lite",
                    "gemini-3.1-flash-lite-preview",
                    "gemini-2.5-flash-lite",
                    "gemini-2.0-flash-lite");
}
