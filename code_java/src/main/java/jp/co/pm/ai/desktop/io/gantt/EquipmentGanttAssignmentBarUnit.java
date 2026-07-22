package jp.co.pm.ai.desktop.io.gantt;

import java.time.LocalDate;
import java.util.List;

/**
 * 設備ガント上の1編集単位（連続バー／同一ギャップセグメントに属する {@code timeline_events} 群）。
 */
public record EquipmentGanttAssignmentBarUnit(
        String barId,
        List<Integer> timelineEventIndices,
        LocalDate date,
        String machine,
        String taskId,
        String eventKind,
        List<EquipmentGanttAssignmentPerson> persons) {

    public EquipmentGanttAssignmentBarUnit {
        timelineEventIndices =
                timelineEventIndices == null ? List.of() : List.copyOf(timelineEventIndices);
        persons = persons == null ? List.of() : List.copyOf(persons);
        barId = barId != null ? barId : "";
        machine = machine != null ? machine : "";
        taskId = taskId != null ? taskId : "";
        eventKind = eventKind != null ? eventKind : "";
    }
}
