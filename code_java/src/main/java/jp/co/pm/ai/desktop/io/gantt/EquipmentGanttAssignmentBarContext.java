package jp.co.pm.ai.desktop.io.gantt;

import java.util.List;
import java.util.Map;
import java.util.Optional;

import jp.co.pm.ai.desktop.io.JsonTableIo;

/** 編集対象バーから skills 照会に使う工程名・機械名を解決する。 */
public final class EquipmentGanttAssignmentBarContext {

    private static final String COL_MACH = "機械名";
    private static final String COL_PROC = "工程名";

    public record ProcessMachine(String processName, String machineName) {}

    private EquipmentGanttAssignmentBarContext() {}

    public static Optional<ProcessMachine> resolve(
            EquipmentGanttAssignmentMetadata metadata,
            JsonTableIo.SheetTable table,
            String barId) {
        if (metadata == null || barId == null || barId.isBlank()) {
            return Optional.empty();
        }
        Optional<Integer> rowIdx = tableRowIndexForBar(metadata, barId);
        if (rowIdx.isPresent() && table != null) {
            Optional<ProcessMachine> fromTable = fromTableRow(table, rowIdx.get());
            if (fromTable.isPresent()) {
                return fromTable;
            }
        }
        return fromBarUnit(metadata, barId);
    }

    private static Optional<Integer> tableRowIndexForBar(
            EquipmentGanttAssignmentMetadata metadata, String barId) {
        for (EquipmentGanttAssignmentSlotBinding b : metadata.slotBindings()) {
            if (barId.equals(b.barId())) {
                return Optional.of(b.tableRowIndex());
            }
        }
        return Optional.empty();
    }

    private static Optional<ProcessMachine> fromTableRow(JsonTableIo.SheetTable table, int rowIdx) {
        List<Map<String, String>> rows = table.rows();
        if (rowIdx < 0 || rowIdx >= rows.size()) {
            return Optional.empty();
        }
        Map<String, String> row = rows.get(rowIdx);
        if (row == null) {
            return Optional.empty();
        }
        String proc = row.getOrDefault(COL_PROC, "").strip();
        String mach = row.getOrDefault(COL_MACH, "").strip();
        if (mach.isEmpty()) {
            return Optional.empty();
        }
        return Optional.of(new ProcessMachine(proc, mach));
    }

    private static Optional<ProcessMachine> fromBarUnit(
            EquipmentGanttAssignmentMetadata metadata, String barId) {
        for (EquipmentGanttAssignmentBarUnit unit : metadata.barUnits()) {
            if (!barId.equals(unit.barId())) {
                continue;
            }
            String[] split =
                    EquipmentGanttContractSheetTableBuilder.splitEquipmentLine(unit.machine());
            String proc = split[0] != null ? split[0].strip() : "";
            String mach = split[1] != null ? split[1].strip() : "";
            if (mach.isEmpty() && unit.machine() != null) {
                mach = unit.machine().strip();
            }
            if (mach.isEmpty()) {
                return Optional.empty();
            }
            return Optional.of(new ProcessMachine(proc, mach));
        }
        return Optional.empty();
    }
}
