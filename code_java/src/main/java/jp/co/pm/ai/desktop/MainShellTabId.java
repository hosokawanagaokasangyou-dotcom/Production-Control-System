package jp.co.pm.ai.desktop;

/**
 * Stable ids for main-shell tabs (persisted in {@link jp.co.pm.ai.desktop.config.DesktopSessionState}).
 */
public enum MainShellTabId {
    RUN("run"),
    ENV("env"),
    MASTER_SUMMARY("masterSummary"),
    PLAN_INPUT("planInput"),
    STAGE1_PREVIEW("stage1Preview"),
    EXCLUDE_RULES("excludeRules"),
    ACTUALS_STATUS("actualsStatus"),
    RESULT_DISPATCH("resultDispatch"),
    PLAN_RESULT_VIEWER("planResultViewer"),
    OPERATOR_CARD("operatorCard"),
    DISPATCH_INTERACTIVE("dispatchInteractive");

    private final String key;

    MainShellTabId(String key) {
        this.key = key;
    }

    public String key() {
        return key;
    }

    public static MainShellTabId fromKey(String k) {
        if (k == null || k.isBlank()) {
            return null;
        }
        String t = k.trim();
        for (MainShellTabId id : values()) {
            if (id.key.equals(t)) {
                return id;
            }
        }
        return null;
    }
}
