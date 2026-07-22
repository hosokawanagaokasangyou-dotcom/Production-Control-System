package jp.co.pm.ai.desktop.io.gantt;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;

/**
 * 担当割当の編集状態（MOVE / SWAP / 追加 / 削除）。純粋モデルで UI から独立。
 */
public final class EquipmentGanttAssignmentEditModel {

    public enum Failure {
        UNKNOWN_BAR,
        UNKNOWN_PERSON,
        DUPLICATE_PERSON,
        EMPTY_BAR_FORBIDDEN,
        SAME_BAR_NOOP
    }

    private final Map<String, List<EquipmentGanttAssignmentPerson>> personsByBarId =
            new LinkedHashMap<>();

    public EquipmentGanttAssignmentEditModel(EquipmentGanttAssignmentMetadata metadata) {
        if (metadata == null) {
            return;
        }
        for (EquipmentGanttAssignmentBarUnit unit : metadata.barUnits()) {
            personsByBarId.put(unit.barId(), new ArrayList<>(unit.persons()));
        }
    }

    public List<EquipmentGanttAssignmentPerson> personsOnBar(String barId) {
        return List.copyOf(personsByBarId.getOrDefault(barId, List.of()));
    }

    public Optional<Failure> movePerson(String fromBarId, String toBarId, String memberKey) {
        if (Objects.equals(fromBarId, toBarId)) {
            return Optional.of(Failure.SAME_BAR_NOOP);
        }
        if (!personsByBarId.containsKey(fromBarId) || !personsByBarId.containsKey(toBarId)) {
            return Optional.of(Failure.UNKNOWN_BAR);
        }
        List<EquipmentGanttAssignmentPerson> from = personsByBarId.get(fromBarId);
        List<EquipmentGanttAssignmentPerson> to = personsByBarId.get(toBarId);
        if (!to.isEmpty()) {
            return swapPerson(fromBarId, toBarId, memberKey, to.getFirst().memberKey());
        }
        EquipmentGanttAssignmentPerson person = removePersonByKey(from, memberKey);
        if (person == null) {
            return Optional.of(Failure.UNKNOWN_PERSON);
        }
        if (from.isEmpty()) {
            from.add(person);
            return Optional.of(Failure.EMPTY_BAR_FORBIDDEN);
        }
        if (containsMemberKey(to, person.memberKey())) {
            from.add(person);
            return Optional.of(Failure.DUPLICATE_PERSON);
        }
        to.add(person.withRole(roleForIndex(to.size())));
        normalizeRoles(from);
        normalizeRoles(to);
        return Optional.empty();
    }

    public Optional<Failure> swapPerson(
            String fromBarId, String toBarId, String fromMemberKey, String toMemberKey) {
        if (Objects.equals(fromBarId, toBarId)) {
            return Optional.of(Failure.SAME_BAR_NOOP);
        }
        if (!personsByBarId.containsKey(fromBarId) || !personsByBarId.containsKey(toBarId)) {
            return Optional.of(Failure.UNKNOWN_BAR);
        }
        List<EquipmentGanttAssignmentPerson> from = personsByBarId.get(fromBarId);
        List<EquipmentGanttAssignmentPerson> to = personsByBarId.get(toBarId);
        EquipmentGanttAssignmentPerson a = removePersonByKey(from, fromMemberKey);
        EquipmentGanttAssignmentPerson b = removePersonByKey(to, toMemberKey);
        if (a == null || b == null) {
            if (a != null) {
                from.add(a);
            }
            if (b != null) {
                to.add(b);
            }
            return Optional.of(Failure.UNKNOWN_PERSON);
        }
        from.add(b.withRole(roleForIndex(from.size())));
        to.add(a.withRole(roleForIndex(to.size())));
        normalizeRoles(from);
        normalizeRoles(to);
        return Optional.empty();
    }

    public Optional<Failure> addPerson(String barId, EquipmentGanttAssignmentPerson person) {
        if (!personsByBarId.containsKey(barId) || person == null) {
            return Optional.of(Failure.UNKNOWN_BAR);
        }
        List<EquipmentGanttAssignmentPerson> list = personsByBarId.get(barId);
        if (containsMemberKey(list, person.memberKey())) {
            return Optional.of(Failure.DUPLICATE_PERSON);
        }
        list.add(person.withRole(roleForIndex(list.size())));
        normalizeRoles(list);
        return Optional.empty();
    }

    public Optional<Failure> removePerson(String barId, String memberKey) {
        if (!personsByBarId.containsKey(barId)) {
            return Optional.of(Failure.UNKNOWN_BAR);
        }
        List<EquipmentGanttAssignmentPerson> list = personsByBarId.get(barId);
        EquipmentGanttAssignmentPerson removed = removePersonByKey(list, memberKey);
        if (removed == null) {
            return Optional.of(Failure.UNKNOWN_PERSON);
        }
        if (list.isEmpty()) {
            list.add(removed);
            return Optional.of(Failure.EMPTY_BAR_FORBIDDEN);
        }
        normalizeRoles(list);
        return Optional.empty();
    }

    /** barId → 編集後の担当者一覧（契約 JSON 反映用）。 */
    public Map<String, List<EquipmentGanttAssignmentPerson>> snapshotPersonsByBarId() {
        Map<String, List<EquipmentGanttAssignmentPerson>> out = new LinkedHashMap<>();
        for (Map.Entry<String, List<EquipmentGanttAssignmentPerson>> e : personsByBarId.entrySet()) {
            out.put(e.getKey(), List.copyOf(e.getValue()));
        }
        return Map.copyOf(out);
    }

    private static EquipmentGanttAssignmentRole roleForIndex(int index) {
        return index == 0
                ? EquipmentGanttAssignmentRole.PRIMARY
                : EquipmentGanttAssignmentRole.SUB;
    }

    private static void normalizeRoles(List<EquipmentGanttAssignmentPerson> list) {
        for (int i = 0; i < list.size(); i++) {
            EquipmentGanttAssignmentPerson p = list.get(i);
            EquipmentGanttAssignmentRole want = roleForIndex(i);
            if (p.role() != want) {
                list.set(i, p.withRole(want));
            }
        }
    }

    private static boolean containsMemberKey(
            List<EquipmentGanttAssignmentPerson> list, String memberKey) {
        for (EquipmentGanttAssignmentPerson p : list) {
            if (p.memberKey().equals(memberKey)) {
                return true;
            }
        }
        return false;
    }

    private static EquipmentGanttAssignmentPerson removePersonByKey(
            List<EquipmentGanttAssignmentPerson> list, String memberKey) {
        for (int i = 0; i < list.size(); i++) {
            if (list.get(i).memberKey().equals(memberKey)) {
                return list.remove(i);
            }
        }
        return null;
    }
}
