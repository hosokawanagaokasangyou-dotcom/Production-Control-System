package jp.co.pm.ai.desktop.io.gantt;

/** 担当割当 DnD 確定時のコールバック（true = 適用成功）。 */
@FunctionalInterface
public interface EquipmentGanttAssignmentDropHandler {
    boolean onDrop(
            EquipmentGanttAssignmentDragPayload source,
            EquipmentGanttAssignmentDropTarget target);
}
