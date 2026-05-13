package jp.co.pm.ai.planning.stage2.core;

import java.util.List;

import jp.co.pm.ai.planning.stage2.Stage2RunContext;
import jp.co.pm.ai.planning.stage2.input.Stage2InputSnapshot;

/**
 * {@link Stage2DispatchLoop} の拡張用フック: マスタシート検出結果をログに出し、Python ログと突き合わせやすくする。
 */
public final class Stage2DispatchDiagnostics {

    private Stage2DispatchDiagnostics() {}

    public static void logMasterProbe(Stage2RunContext ctx, Stage2InputSnapshot snap) {
        try {
            Stage2MasterSheetProbe p = Stage2MasterSheetProbe.scan(snap.masterPath());
            ctx.log(
                    "[stage2-java] master_probe: sheets="
                            + p.sheetCount()
                            + " need="
                            + p.hasNeedSheet()
                            + " machine_calendar="
                            + p.hasMachineCalendarSheet());
        } catch (Exception e) {
            ctx.log("[stage2-java] master_probe: 失敗 " + e.getMessage());
        }
    }

    public static void logTaskQueuePreview(Stage2RunContext ctx, List<Stage2QueuedTask> queue, int maxIds) {
        if (queue == null || queue.isEmpty()) {
            ctx.log("[stage2-java] task_queue_preview: (empty)");
            return;
        }
        int lim = Math.min(queue.size(), Math.max(1, maxIds));
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < lim; i++) {
            if (i > 0) {
                sb.append(',');
            }
            sb.append(queue.get(i).requestId());
        }
        if (queue.size() > lim) {
            sb.append(",…(total ").append(queue.size()).append(')');
        }
        ctx.log("[stage2-java] task_queue_preview: " + sb);
    }
}
