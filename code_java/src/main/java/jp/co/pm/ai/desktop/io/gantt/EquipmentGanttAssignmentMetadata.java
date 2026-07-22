package jp.co.pm.ai.desktop.io.gantt;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/** 契約 JSON から構築した担当割当編集メタデータ一式。 */
public record EquipmentGanttAssignmentMetadata(
        List<EquipmentGanttAssignmentBarUnit> barUnits,
        List<EquipmentGanttAssignmentSlotBinding> slotBindings) {

    public EquipmentGanttAssignmentMetadata {
        barUnits = barUnits == null ? List.of() : List.copyOf(barUnits);
        slotBindings = slotBindings == null ? List.of() : List.copyOf(slotBindings);
    }

    public static EquipmentGanttAssignmentMetadata empty() {
        return new EquipmentGanttAssignmentMetadata(List.of(), List.of());
    }

    public Map<String, EquipmentGanttAssignmentBarUnit> barUnitsById() {
        Map<String, EquipmentGanttAssignmentBarUnit> out = new LinkedHashMap<>();
        for (EquipmentGanttAssignmentBarUnit u : barUnits) {
            if (u != null && !u.barId().isBlank()) {
                out.put(u.barId(), u);
            }
        }
        return Map.copyOf(out);
    }

    /** 表行とスロット範囲から編集単位 barId を解決する。 */
    public java.util.Optional<String> resolveBarId(int tableRowIndex, int fromSlot, int toSlot) {
        for (EquipmentGanttAssignmentSlotBinding b : slotBindings) {
            if (b.tableRowIndex() != tableRowIndex) {
                continue;
            }
            if (b.fromSlot() <= toSlot && b.toSlot() >= fromSlot) {
                return java.util.Optional.of(b.barId());
            }
        }
        return java.util.Optional.empty();
    }

    /**
     * UI のバー run に対応する barId。スロット binding が無い／ずれているときはタイムライン文言・担当者・機械名で補完する。
     */
    public java.util.Optional<String> resolveBarIdForBadgeRun(
            int tableRowIndex,
            int fromSlot,
            int toSlot,
            String machineDisplay,
            List<String> badgePersonLabels,
            String runTimelineText) {
        java.util.Optional<String> hit = resolveBarId(tableRowIndex, fromSlot, toSlot);
        if (hit.isPresent()) {
            return hit;
        }
        for (int s = fromSlot; s <= toSlot; s++) {
            java.util.Optional<String> single = resolveBarId(tableRowIndex, s, s);
            if (single.isPresent()) {
                return single;
            }
        }
        hit =
                resolveBarIdByRunText(
                        tableRowIndex,
                        fromSlot,
                        toSlot,
                        machineDisplay,
                        runTimelineText,
                        badgePersonLabels);
        if (hit.isPresent()) {
            return hit;
        }
        return resolveBarIdByPersonLabels(tableRowIndex, machineDisplay, badgePersonLabels);
    }

    /** @deprecated {@link #resolveBarIdForBadgeRun(int, int, int, String, List, String)} を使用 */
    @Deprecated
    public java.util.Optional<String> resolveBarIdForBadgeRun(
            int tableRowIndex,
            int fromSlot,
            int toSlot,
            String machineDisplay,
            List<String> badgePersonLabels) {
        return resolveBarIdForBadgeRun(
                tableRowIndex, fromSlot, toSlot, machineDisplay, badgePersonLabels, "");
    }

    private java.util.Optional<String> resolveBarIdByRunText(
            int tableRowIndex,
            int fromSlot,
            int toSlot,
            String machineDisplay,
            String runTimelineText,
            List<String> badgePersonLabels) {
        String text = runTimelineText != null ? runTimelineText.strip() : "";
        if (text.isEmpty()) {
            return java.util.Optional.empty();
        }
        String taskKey = EquipmentGanttPersonBadgeRunMerge.mergeKey(text);
        List<String> candidates = new ArrayList<>();
        for (EquipmentGanttAssignmentBarUnit unit : barUnits) {
            if (!machineDisplayMatches(unit.machine(), machineDisplay)) {
                continue;
            }
            if (!taskKeyMatchesUnit(taskKey, unit)) {
                continue;
            }
            candidates.add(unit.barId());
        }
        if (candidates.isEmpty()) {
            return java.util.Optional.empty();
        }
        if (candidates.size() > 1
                && badgePersonLabels != null
                && !badgePersonLabels.isEmpty()) {
            List<String> byPerson = new ArrayList<>();
            for (EquipmentGanttAssignmentBarUnit unit : barUnits) {
                if (!candidates.contains(unit.barId())) {
                    continue;
                }
                if (personLabelsMatchUnit(unit, badgePersonLabels)) {
                    byPerson.add(unit.barId());
                }
            }
            if (!byPerson.isEmpty()) {
                candidates = byPerson;
            }
        }
        if (candidates.size() == 1) {
            return java.util.Optional.of(candidates.getFirst());
        }
        for (String barId : candidates) {
            for (EquipmentGanttAssignmentSlotBinding b : slotBindings) {
                if (b.tableRowIndex() != tableRowIndex || !barId.equals(b.barId())) {
                    continue;
                }
                if (b.fromSlot() <= toSlot && b.toSlot() >= fromSlot) {
                    return java.util.Optional.of(barId);
                }
            }
        }
        for (String barId : candidates) {
            for (EquipmentGanttAssignmentSlotBinding b : slotBindings) {
                if (b.tableRowIndex() == tableRowIndex && barId.equals(b.barId())) {
                    return java.util.Optional.of(barId);
                }
            }
        }
        return java.util.Optional.of(candidates.getFirst());
    }

    private java.util.Optional<String> resolveBarIdByPersonLabels(
            int tableRowIndex, String machineDisplay, List<String> badgePersonLabels) {
        if (badgePersonLabels == null || badgePersonLabels.isEmpty()) {
            return java.util.Optional.empty();
        }
        for (String rawLabel : badgePersonLabels) {
            if (rawLabel == null || rawLabel.isBlank()) {
                continue;
            }
            String label = rawLabel.strip();
            for (EquipmentGanttAssignmentBarUnit unit : barUnits) {
                if (!machineDisplayMatches(unit.machine(), machineDisplay)) {
                    continue;
                }
                if (!unitEligibleOnTableRow(unit, tableRowIndex)) {
                    continue;
                }
                for (EquipmentGanttAssignmentPerson p : unit.persons()) {
                    if (personLabelMatchesBadge(p, label)) {
                        return java.util.Optional.of(unit.barId());
                    }
                }
            }
        }
        return java.util.Optional.empty();
    }

    private boolean unitEligibleOnTableRow(EquipmentGanttAssignmentBarUnit unit, int tableRowIndex) {
        boolean unitHasBindings =
                slotBindings.stream().anyMatch(b -> unit.barId().equals(b.barId()));
        if (!unitHasBindings) {
            return true;
        }
        return slotBindings.stream()
                .anyMatch(
                        b ->
                                b.tableRowIndex() == tableRowIndex
                                        && unit.barId().equals(b.barId()));
    }

    private static boolean personLabelsMatchUnit(
            EquipmentGanttAssignmentBarUnit unit, List<String> badgePersonLabels) {
        if (badgePersonLabels == null || badgePersonLabels.isEmpty()) {
            return true;
        }
        for (String raw : badgePersonLabels) {
            if (raw == null || raw.isBlank()) {
                continue;
            }
            String label = raw.strip();
            for (EquipmentGanttAssignmentPerson p : unit.persons()) {
                if (personLabelMatchesBadge(p, label)) {
                    return true;
                }
            }
        }
        return false;
    }

    public static boolean personLabelMatchesBadge(EquipmentGanttAssignmentPerson p, String label) {
        if (p == null || label == null || label.isBlank()) {
            return false;
        }
        String l = label.strip();
        if (l.equals(p.badgeLabel()) || l.equals(p.fullName())) {
            return true;
        }
        if (!p.badgeLabel().isBlank() && l.contains(p.badgeLabel())) {
            return true;
        }
        String badgeFromFull = PersonNameBadgeText.badgeTwoFromRawName(p.fullName());
        if (!badgeFromFull.isBlank() && l.equals(badgeFromFull)) {
            return true;
        }
        String sei = PersonNameBadgeText.surnameLabelOnly(p.fullName());
        return !sei.isBlank() && (l.equals(sei) || sei.startsWith(l) || l.startsWith(sei));
    }

    private static boolean machineDisplayMatches(String eventMachine, String rowMachineDisplay) {
        return EquipmentGanttContractSheetTableBuilder.equipmentColumnMatchesEventMachine(
                rowMachineDisplay, eventMachine);
    }

    private static boolean taskKeyMatchesUnit(String runMergeKey, EquipmentGanttAssignmentBarUnit unit) {
        String kindPrefix = "DEFAULT";
        String identity = runMergeKey != null ? runMergeKey : "";
        int sep = identity.indexOf('\u0001');
        if (sep >= 0) {
            kindPrefix = identity.substring(0, sep);
            identity = identity.substring(sep + 1);
        }
        if (!"DEFAULT".equals(kindPrefix)
                && slotKindMatchesEventKind(kindPrefix, unit.eventKind())) {
            return true;
        }
        if (identity.isBlank()) {
            return true;
        }
        String tid = unit.taskId() != null ? unit.taskId().strip() : "";
        if (!tid.isEmpty()) {
            if (identity.contains(tid)) {
                return true;
            }
            String core = identity.replaceAll("\\s+休憩[前後](\\s+.*)?$", "").strip();
            if (!core.isEmpty()
                    && (core.contains(tid) || tid.contains(core) || core.startsWith(tid))) {
                return true;
            }
        }
        return identity.equals(unit.eventKind());
    }

    public static boolean slotKindMatchesEventKind(String slotKind, String eventKind) {
        if (slotKind == null || eventKind == null) {
            return false;
        }
        return switch (slotKind) {
            case "STARTUP" -> "machine_daily_startup".equals(eventKind);
            case "REQUEST_SWITCH_PREP" -> "request_switch_prep".equals(eventKind);
            case "BREAK_RESUME_PREP" -> "break_resume_prep".equals(eventKind);
            case "POST_MACHINING_CLEANUP" -> "post_machining_cleanup".equals(eventKind);
            case "REQUEST_INTERVAL_BUFFER" -> "request_interval_buffer".equals(eventKind);
            default -> false;
        };
    }
}
