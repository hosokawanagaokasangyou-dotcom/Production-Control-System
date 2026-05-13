package jp.co.pm.ai.desktop;

/**
 * Stable ids for main-shell tabs (persisted in {@link jp.co.pm.ai.desktop.config.DesktopSessionState}).
 */
public enum MainShellTabId {
    RUN("run"),
    UI_BADGE_DESIGN("uiBadgeDesign"),
    PUSH_BUTTON_DESIGN("pushButtonDesign"),
    ENV("env"),
    /** JVM ヒープ・メモリ監視・次回起動時ヒープ希望値。 */
    MEMORY_SETTINGS("memorySettings"),
    /** UI 全体の既定リセット・パッケージ既定の書き出し。 */
    GLOBAL_SETTINGS("globalSettings"),
    /** ユーザープロファイル（UI 設定の保存・読み出し、{@code ~/.pm-ai-desktop/user-profiles}）。 */
    USER_PROFILES("userProfiles"),
    MASTER_SUMMARY("masterSummary"),
    PLAN_INPUT("planInput"),
    STAGE1_PREVIEW("stage1Preview"),
    EXCLUDE_RULES("excludeRules"),
    SPECIAL_RULES("specialRules"),
    ACTUALS_STATUS("actualsStatus"),
    /** 納期管理（アラジン計画）風ビュー（計画＋実績・計画比較表）。 */
    DELIVERY_CALENDAR_VIEW("deliveryCalendarView"),
    RESULT_DISPATCH("resultDispatch"),
    PLAN_RESULT_VIEWER("planResultViewer"),
    EQUIPMENT_GANTT_GRAPHIC("equipmentGanttGraphic"),
    GANTT_PERSON_BADGE_DESIGN("ganttPersonBadgeDesign"),
    OPERATOR_CARD("operatorCard"),
    DISPATCH_INTERACTIVE("dispatchInteractive"),
    /** 配台ワークスペースのスナップショット履歴（結果 JSON・ガント表示・列順の復元）。 */
    PLAN_WORKSPACE_HISTORY("planWorkspaceHistory"),
    /** Gemini generateContent の往復レイテンシ計測。 */
    API_MODEL_BENCHMARK("apiModelBenchmark"),
    /** メインシェル末尾の「タブ整理」（入れ子構成・色の編集用）。 */
    TAB_ORGANIZER("tabOrganizer");

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
