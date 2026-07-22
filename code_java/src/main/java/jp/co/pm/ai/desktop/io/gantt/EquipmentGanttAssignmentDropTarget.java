package jp.co.pm.ai.desktop.io.gantt;

/** ドロップ先（バー本体または特定バッジ）。memberKey が空ならバー本体への MOVE。 */
public record EquipmentGanttAssignmentDropTarget(String barId, String memberKey) {

    public EquipmentGanttAssignmentDropTarget {
        barId = barId != null ? barId : "";
        memberKey = memberKey != null ? memberKey : "";
    }
}
