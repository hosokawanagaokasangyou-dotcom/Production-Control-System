package jp.co.pm.ai.desktop.io.gantt;

import java.util.ArrayList;
import java.util.List;

import jp.co.pm.ai.desktop.ui.GanttScheduleSlotBarKind;

/**
 * 設備ガントのタイムラインスロット列からバー run を切り出す（{@link
 * jp.co.pm.ai.desktop.ui.EquipmentGraphicGanttPane#collectBarRunsForPersonBadges} と同一規則）。
 */
public final class EquipmentGanttPersonBadgeRunMerge {

    private EquipmentGanttPersonBadgeRunMerge() {}

    public static String mergeKey(String slotText) {
        String t = slotText != null ? slotText.strip() : "";
        if (t.isEmpty()) {
            return "";
        }
        GanttScheduleSlotBarKind k = GanttScheduleSlotBarKind.fromTimelineCell(t);
        if (k == GanttScheduleSlotBarKind.BREAK
                || k == GanttScheduleSlotBarKind.STARTUP
                || k == GanttScheduleSlotBarKind.REQUEST_SWITCH_PREP
                || k == GanttScheduleSlotBarKind.BREAK_RESUME_PREP
                || k == GanttScheduleSlotBarKind.POST_MACHINING_CLEANUP
                || k == GanttScheduleSlotBarKind.REQUEST_INTERVAL_BUFFER) {
            return k.name() + "\u0001" + t;
        }
        String base = t.replaceFirst("\\s+\\d+(?:\\.\\d+)?m\\s*$", "").strip();
        String identity = base.isEmpty() ? t : base;
        return k.name() + "\u0001" + identity;
    }

    public static List<EquipmentGanttAssignmentMetadataBuilder.BarRunSlice> collectRuns(
            List<String> slotTexts) {
        int n = slotTexts != null ? slotTexts.size() : 0;
        List<EquipmentGanttAssignmentMetadataBuilder.BarRunSlice> runs = new ArrayList<>();
        int runStart = -1;
        String runKey = "";
        String headText = "";
        for (int i = 0; i < n; i++) {
            String t = slotTexts.get(i) != null ? slotTexts.get(i).strip() : "";
            boolean empty = t.isEmpty();
            if (empty) {
                if (runStart >= 0) {
                    runs.add(
                            new EquipmentGanttAssignmentMetadataBuilder.BarRunSlice(
                                    runStart, i - 1, headText));
                    runStart = -1;
                    runKey = "";
                    headText = "";
                }
                continue;
            }
            String key = mergeKey(t);
            if (runStart < 0) {
                runStart = i;
                runKey = key;
                headText = t;
            } else if (!key.equals(runKey)) {
                runs.add(
                        new EquipmentGanttAssignmentMetadataBuilder.BarRunSlice(
                                runStart, i - 1, headText));
                runStart = i;
                runKey = key;
                headText = t;
            }
        }
        if (runStart >= 0) {
            runs.add(
                    new EquipmentGanttAssignmentMetadataBuilder.BarRunSlice(
                            runStart, n - 1, headText));
        }
        return runs;
    }
}
