package jp.co.pm.ai.planning.stage2.core;

/**
 * Python 正本 {@code planning_core/_core.py} の行帯索引（{@code planning_core/_core_FILE_MAP.txt}）を Java 側から参照するための定数。
 * 移植作業のナビゲーション用であり、行番号はソース変更でずれることに注意。
 */
public final class Stage2PlanningCorePyBands {

    private Stage2PlanningCorePyBands() {}

    /** 配台用タスクキュー（DataFrame 行の除外・dict 化の始まり）。 */
    public static final String BUILD_TASK_QUEUE_FROM_PLANNING_DF =
            "planning_core/_core.py: build_task_queue_from_planning_df（目安 9988 行付近）";

    /** 日内配台の試行順 1 パス。 */
    public static final String TRIAL_ORDER_FIRST_SCHEDULE_PASS =
            "planning_core/_core.py: _trial_order_first_schedule_pass（目安 20576 行付近）";

    /** 段階2メイン（成果物・メンバー出力を含む）。 */
    public static final String GENERATE_PLAN_IMPL =
            "planning_core/_core.py: _generate_plan_impl（目安 21795 行付近）";
}
