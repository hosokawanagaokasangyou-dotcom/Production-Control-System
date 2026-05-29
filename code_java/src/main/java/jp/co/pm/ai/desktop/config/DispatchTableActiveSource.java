package jp.co.pm.ai.desktop.config;

/** 配台手動修正・段階3 等が参照する結果_配台表 JSON の正本（段階2 出力か段階2.5 整列後か）。 */
public enum DispatchTableActiveSource {
    STAGE2("stage2", "段階2"),
    STAGE2_5("stage2_5", "段階2.5(AI)");

    private final String envToken;
    private final String displayLabel;

    DispatchTableActiveSource(String envToken, String displayLabel) {
        this.envToken = envToken;
        this.displayLabel = displayLabel;
    }

    public String envToken() {
        return envToken;
    }

    public String displayLabel() {
        return displayLabel;
    }

    public static DispatchTableActiveSource fromEnvToken(String raw) {
        if (raw != null && STAGE2_5.envToken.equalsIgnoreCase(raw.strip())) {
            return STAGE2_5;
        }
        return STAGE2;
    }
}
