package jp.co.pm.ai.desktop.io.gantt;

import java.util.Map;

/** 設備ガント build へ渡す担当割当編集インタラクション。 */
public record EquipmentGanttAssignmentInteraction(
        boolean enabled,
        EquipmentGanttAssignmentMetadata metadata,
        Map<String, java.util.List<EquipmentGanttAssignmentPerson>> personsByBarId,
        EquipmentGanttAssignmentDropHandler dropHandler,
        EquipmentGanttAssignmentEditActions editActions) {

    public static EquipmentGanttAssignmentInteraction disabled() {
        return new EquipmentGanttAssignmentInteraction(
                false, EquipmentGanttAssignmentMetadata.empty(), Map.of(), null, null);
    }

    public boolean active() {
        return enabled && dropHandler != null && metadata != null && !metadata.barUnits().isEmpty();
    }

    public boolean editMenuActive() {
        return active() && editActions != null;
    }
}
