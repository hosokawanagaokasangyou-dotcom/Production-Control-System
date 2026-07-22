package jp.co.pm.ai.desktop.io.gantt;

/** 表行・スロット範囲と編集単位 {@link EquipmentGanttAssignmentBarUnit#barId()} の対応。 */
public record EquipmentGanttAssignmentSlotBinding(
        int tableRowIndex, int fromSlot, int toSlot, String barId) {

    public EquipmentGanttAssignmentSlotBinding {
        barId = barId != null ? barId : "";
    }
}
