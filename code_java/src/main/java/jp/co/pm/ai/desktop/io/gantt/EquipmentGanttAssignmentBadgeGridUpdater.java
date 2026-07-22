package jp.co.pm.ai.desktop.io.gantt;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/** 担当割当編集結果を担当バッジ行グリッドへ反映する。 */
public final class EquipmentGanttAssignmentBadgeGridUpdater {

    private EquipmentGanttAssignmentBadgeGridUpdater() {}

    public static void applyToBadgeRows(
            List<List<String>> badgeRows,
            EquipmentGanttAssignmentMetadata metadata,
            Map<String, List<EquipmentGanttAssignmentPerson>> personsByBarId) {
        if (badgeRows == null || metadata == null || personsByBarId == null) {
            return;
        }
        for (EquipmentGanttAssignmentSlotBinding binding : metadata.slotBindings()) {
            if (binding.tableRowIndex() < 0 || binding.tableRowIndex() >= badgeRows.size()) {
                continue;
            }
            List<EquipmentGanttAssignmentPerson> persons = personsByBarId.get(binding.barId());
            if (persons == null) {
                continue;
            }
            List<String> labels = new ArrayList<>();
            for (EquipmentGanttAssignmentPerson p : persons) {
                if (p != null && !p.badgeLabel().isBlank()) {
                    labels.add(p.badgeLabel());
                }
            }
            String cell = PersonNameBadgeText.joinBadgeCells(labels);
            List<String> row = badgeRows.get(binding.tableRowIndex());
            if (row == null) {
                continue;
            }
            int to = Math.min(binding.toSlot(), row.size() - 1);
            for (int s = binding.fromSlot(); s <= to; s++) {
                row.set(s, cell);
            }
        }
    }
}
