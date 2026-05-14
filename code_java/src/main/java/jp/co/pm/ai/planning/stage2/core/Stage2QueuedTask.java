package jp.co.pm.ai.planning.stage2.core;

import java.util.Optional;

/**
 * Python {@code build_task_queue_from_planning_df} 相当の行単位表現（段階移植用）。
 *
 * @param excelRowIndex 計画入力表の 1 始まり行番号（ヘッダー=1、最初のデータ=2）
 * @param dispatchTrialOrderFromSheet 「配台試行順番」列に数値があるときのみ。欠損は empty。
 * @param dispatchTrialOrderEffective {@link Stage2DispatchTrialOrderApplier} 適用後の試行順（Python
 *     {@code dispatch_trial_order}）。未適用は 0。
 */
public record Stage2QueuedTask(
        int excelRowIndex,
        String requestId,
        Optional<Integer> dispatchTrialOrderFromSheet,
        int dispatchTrialOrderEffective) {

    /** 列「配台試行順番」未読込・未適用の最小コンストラクタ（後方互換）。 */
    public Stage2QueuedTask(int excelRowIndex, String requestId) {
        this(excelRowIndex, requestId, Optional.empty(), 0);
    }
}
