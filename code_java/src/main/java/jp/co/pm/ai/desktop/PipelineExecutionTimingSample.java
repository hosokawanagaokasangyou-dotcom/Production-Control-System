package jp.co.pm.ai.desktop;

/** パイプライン実行時間の1計測サンプル。 */
public record PipelineExecutionTimingSample(
        PipelineExecutionTimingKind kind,
        long finishedAtEpochMs,
        long durationMs,
        String writerHost,
        String writerIp) {

    public PipelineExecutionTimingSample(
            PipelineExecutionTimingKind kind, long finishedAtEpochMs, long durationMs) {
        this(kind, finishedAtEpochMs, durationMs, "", "");
    }

    public String writerEndpointLabel() {
        if (writerIp != null && !writerIp.isBlank()) {
            if (writerHost != null && !writerHost.isBlank()) {
                return writerHost + " (" + writerIp + ")";
            }
            return writerIp;
        }
        if (writerHost != null && !writerHost.isBlank()) {
            return writerHost;
        }
        return "—";
    }
}
