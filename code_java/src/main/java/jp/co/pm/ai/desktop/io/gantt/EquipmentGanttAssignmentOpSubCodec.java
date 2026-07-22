package jp.co.pm.ai.desktop.io.gantt;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/** 編集後の担当者一覧を契約 JSON の {@code op} / {@code sub} 文字列へ変換する。 */
public final class EquipmentGanttAssignmentOpSubCodec {

    private EquipmentGanttAssignmentOpSubCodec() {}

    public static OpSubPair encode(List<EquipmentGanttAssignmentPerson> persons) {
        if (persons == null || persons.isEmpty()) {
            return new OpSubPair("", "");
        }
        String op = persons.get(0).fullName();
        if (persons.size() == 1) {
            return new OpSubPair(op, "");
        }
        List<String> subs = new ArrayList<>();
        for (int i = 1; i < persons.size(); i++) {
            subs.add(persons.get(i).fullName());
        }
        return new OpSubPair(op, String.join(",", subs));
    }

    /**
     * 編集モデルのスナップショットを {@code timeline_events} の index 群へ反映した op/sub マップを返す。
     * キーはイベント index、値は更新後の op/sub。
     */
    public static Map<Integer, OpSubPair> eventUpdates(
            EquipmentGanttAssignmentMetadata metadata,
            Map<String, List<EquipmentGanttAssignmentPerson>> personsByBarId) {
        if (metadata == null || personsByBarId == null) {
            return Map.of();
        }
        Map<Integer, OpSubPair> out = new java.util.LinkedHashMap<>();
        for (EquipmentGanttAssignmentBarUnit unit : metadata.barUnits()) {
            List<EquipmentGanttAssignmentPerson> persons =
                    personsByBarId.getOrDefault(unit.barId(), unit.persons());
            OpSubPair pair = encode(persons);
            for (int idx : unit.timelineEventIndices()) {
                out.put(idx, pair);
            }
        }
        return Map.copyOf(out);
    }

    public record OpSubPair(String op, String sub) {}
}
