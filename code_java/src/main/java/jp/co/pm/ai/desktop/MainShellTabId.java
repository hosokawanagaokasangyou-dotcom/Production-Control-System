package jp.co.pm.ai.desktop;

/**
 * Stable ids for main-shell tabs (persisted in {@link jp.co.pm.ai.desktop.config.DesktopSessionState}).
 */
public enum MainShellTabId {
    /** 加工実績・予定の機械別ダッシュボード（実行・ログの左＝先頭）。 */
    EQUIPMENT_STATUS_DASHBOARD("equipmentStatusDashboard"),
    /** 加工実績・加工予定を日別に重ねたトレンド（日次棒 + 累計折れ線）。 */
    PROCESSING_TREND("processingTrend"),
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
    /** 工場別の配台システム操作者名（起動時選択・作成者表示）。 */
    OPERATOR_USER_MANAGEMENT("operatorUserManagement"),
    /** 会社カレンダー（公休・特別休暇・出勤日）。 */
    COMPANY_CALENDAR("companyCalendar"),
    /** メンバー勤怠（カレンダー方式）。 */
    MEMBER_ATTENDANCE("memberAttendance"),
    /** 機械カレンダー（JSON 正本）。 */
    MACHINE_CALENDAR("machineCalendar"),
    /** MASTER（skills / need / speed / 組み合わせ表の JSON）。 */
    MASTER_DISPATCH_SHEETS("masterDispatchSheets"),
    MASTER_SUMMARY("masterSummary"),
    PLAN_INPUT("planInput"),
    /** 加工依頼書の照合・対比型入力（湖南工場・ReconciliationApp 由来）。 */
    REQUEST_FORM_INPUT("requestFormInput"),
    /** 依頼書原本の受注転記率・アラジン加工計画の確認。 */
    REQUEST_FORM_PIPELINE_CHECK("requestFormPipelineCheck"),
    /** RDP 接続・RAP 設定・接続先ランチャー配備。 */
    REMOTE_DESKTOP("remoteDesktop"),
    STAGE1_PREVIEW("stage1Preview"),
    /** メイン画面「材料・製品種類情報」: {@code code/} 配下の製品・原反キー・値テーブル。 */
    CODE_LOOKUP_TABLES("codeLookupTables"),
    EXCLUDE_RULES("excludeRules"),
    SPECIAL_RULES("specialRules"),
    ACTUALS_STATUS("actualsStatus"),
    /** 加工日報発行問合せ CSV の表形式閲覧。 */
    DAILY_REPORT_CSV_VIEW("dailyReportCsvView"),
    /** 納期管理（アラジン計画）風ビュー（計画＋実績・計画比較表）。 */
    DELIVERY_CALENDAR_VIEW("deliveryCalendarView"),
    RESULT_DISPATCH("resultDispatch"),
    PLAN_RESULT_VIEWER("planResultViewer"),
    EQUIPMENT_GANTT_GRAPHIC("equipmentGanttGraphic"),
    GANTT_PERSON_BADGE_DESIGN("ganttPersonBadgeDesign"),
    /** 依頼書プレビュー・原本更新バッジのデザイン。 */
    REQUEST_FORM_PREVIEW_BADGE_DESIGN("requestFormPreviewBadgeDesign"),
    OPERATOR_CARD("operatorCard"),
    /** 配台ワークスペースのスナップショット履歴（結果 JSON・ガント表示・列順の復元）。 */
    PLAN_WORKSPACE_HISTORY("planWorkspaceHistory"),
    /** 段階1キャッシュ等の退避履歴（クリア前退避・復元）。 */
    CACHE_HISTORY("cacheHistory"),
    /** 配台重要操作の操作者別ログ（共有フォルダ、90日）。 */
    OPERATOR_ACTION_LOG("operatorActionLog"),
    /** 同一化チェック結果の操作者別履歴（Excel＋加工計画 JSON、最新20件）。 */
    IDENTITY_CHECK_HISTORY("identityCheckHistory"),
    /** Gemini generateContent の往復レイテンシ計測。 */
    API_MODEL_BENCHMARK("apiModelBenchmark"),
    /** 段階1～3・サマリ Excel・納期管理ビューの実行時間トレンド。 */
    PIPELINE_EXECUTION_TIMING("pipelineExecutionTiming"),
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
