package jp.co.pm.ai.desktop.io.gantt;

import java.util.Map;
import java.util.function.BiPredicate;

/** 設備ガント担当割当編集のドラッグ＆ドロップ用ペイロード。 */
public record EquipmentGanttAssignmentDragPayload(String barId, String memberKey) {

    public static final String DRAG_PREFIX = "pm-ai-gantt-assignment:";

    public EquipmentGanttAssignmentDragPayload {
        barId = barId != null ? barId : "";
        memberKey = memberKey != null ? memberKey : "";
    }

    public String encode() {
        return DRAG_PREFIX + barId + '\u001f' + memberKey;
    }

    public static EquipmentGanttAssignmentDragPayload decode(String raw) {
        if (raw == null || !raw.startsWith(DRAG_PREFIX)) {
            return null;
        }
        String body = raw.substring(DRAG_PREFIX.length());
        int sep = body.indexOf('\u001f');
        if (sep < 0) {
            return new EquipmentGanttAssignmentDragPayload(body, "");
        }
        return new EquipmentGanttAssignmentDragPayload(
                body.substring(0, sep), body.substring(sep + 1));
    }
}
