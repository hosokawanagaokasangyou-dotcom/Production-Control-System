package jp.co.pm.ai.desktop;

/** パイプライン実行時間の1計測サンプル。 */
public record PipelineExecutionTimingSample(
        PipelineExecutionTimingKind kind, long finishedAtEpochMs, long durationMs) {}
