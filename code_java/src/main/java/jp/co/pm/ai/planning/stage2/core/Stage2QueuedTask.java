package jp.co.pm.ai.planning.stage2.core;

/**
 * Python {@code build_task_queue_from_planning_df} 相当の行単位表現（段階移植用の最小フィールド）。
 *
 * @param excelRowIndex 計画入力表の 1 始まり行番号（ヘッダー=1、最初のデータ=2）
 */
public record Stage2QueuedTask(int excelRowIndex, String requestId) {}
