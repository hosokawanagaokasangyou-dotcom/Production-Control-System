package jp.co.pm.ai.desktop.config;

/** 段階2.5 の学習データ扱い: 蓄積するか、既存のみ参照して推論するか。 */
public enum DispatchStage25LearningMode {
    ACCUMULATE("accumulate", "学習を蓄積"),
    INFERENCE_ONLY("inference_only", "学習推論のみ（蓄積しない）");

    private final String envToken;
    private final String displayLabel;

    DispatchStage25LearningMode(String envToken, String displayLabel) {
        this.envToken = envToken;
        this.displayLabel = displayLabel;
    }

    public String envToken() {
        return envToken;
    }

    public String displayLabel() {
        return displayLabel;
    }

    public static DispatchStage25LearningMode fromEnvToken(String raw) {
        if (raw != null && INFERENCE_ONLY.envToken.equalsIgnoreCase(raw.strip())) {
            return INFERENCE_ONLY;
        }
        return ACCUMULATE;
    }
}
