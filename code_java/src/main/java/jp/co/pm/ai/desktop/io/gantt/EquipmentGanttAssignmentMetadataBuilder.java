package jp.co.pm.ai.desktop.io.gantt;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import jp.co.pm.ai.desktop.io.JsonTableIo;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttContractSheetTableBuilder.TimelineEvent;

/**
 * 設備ガント契約の {@code timeline_events} と表行から、担当割当編集用メタデータを構築する。
 */
public final class EquipmentGanttAssignmentMetadataBuilder {

    private static final String COL_DATE = "日付";
    private static final String COL_MACH = "機械名";

    private EquipmentGanttAssignmentMetadataBuilder() {}

    public static EquipmentGanttAssignmentMetadata build(
            List<TimelineEvent> timelineEvents,
            JsonTableIo.SheetTable table,
            List<List<String>> badgeSlotRows,
            List<LocalTime> slotStarts) {
        if (timelineEvents == null || timelineEvents.isEmpty()) {
            return EquipmentGanttAssignmentMetadata.empty();
        }
        List<EquipmentGanttAssignmentBarUnit> barUnits = buildBarUnits(timelineEvents);
        Map<String, EquipmentGanttAssignmentBarUnit> byId =
                new LinkedHashMap<>();
        for (EquipmentGanttAssignmentBarUnit u : barUnits) {
            byId.put(u.barId(), u);
        }
        List<EquipmentGanttAssignmentSlotBinding> bindings =
                buildSlotBindings(table, badgeSlotRows, slotStarts, timelineEvents, byId);
        return new EquipmentGanttAssignmentMetadata(barUnits, bindings);
    }

    static List<EquipmentGanttAssignmentBarUnit> buildBarUnits(List<TimelineEvent> events) {
        Map<String, List<Integer>> machiningSegments = machiningSegmentGroups(events);
        List<EquipmentGanttAssignmentBarUnit> units = new ArrayList<>();
        for (Map.Entry<String, List<Integer>> en : machiningSegments.entrySet()) {
            List<Integer> indices = en.getValue();
            if (indices.isEmpty()) {
                continue;
            }
            TimelineEvent head = events.get(indices.getFirst());
            units.add(
                    new EquipmentGanttAssignmentBarUnit(
                            barIdForIndices(indices),
                            indices,
                            head.date,
                            head.machine,
                            head.taskId,
                            head.eventKind,
                            personsFromEvent(head)));
        }
        for (int i = 0; i < events.size(); i++) {
            if (machiningIndexCovered(machiningSegments, i)) {
                continue;
            }
            TimelineEvent e = events.get(i);
            if (!hasAssignableOperators(e)) {
                continue;
            }
            units.add(
                    new EquipmentGanttAssignmentBarUnit(
                            barIdForIndices(List.of(i)),
                            List.of(i),
                            e.date,
                            e.machine,
                            e.taskId,
                            e.eventKind,
                            personsFromEvent(e)));
        }
        return List.copyOf(units);
    }

    private static boolean machiningIndexCovered(
            Map<String, List<Integer>> machiningSegments, int index) {
        for (List<Integer> ix : machiningSegments.values()) {
            if (ix.contains(index)) {
                return true;
            }
        }
        return false;
    }

    /**
     * {@link EquipmentGanttContractSheetTableBuilder#applyGapAwareMachiningLabels} と同じセグメント分割で
     * 加工イベント index を束ねる。
     */
    static Map<String, List<Integer>> machiningSegmentGroups(List<TimelineEvent> events) {
        Map<String, List<Integer>> group = new LinkedHashMap<>();
        for (int i = 0; i < events.size(); i++) {
            TimelineEvent e = events.get(i);
            if (!TimelineEvent.isMachiningDispatch(e)) {
                continue;
            }
            group.computeIfAbsent(TimelineEvent.gapGroupKey(e), k -> new ArrayList<>()).add(i);
        }
        Map<String, List<Integer>> out = new LinkedHashMap<>();
        for (Map.Entry<String, List<Integer>> en : group.entrySet()) {
            List<Integer> ix = en.getValue();
            ix.sort(Comparator.comparing(i -> events.get(i).start));
            List<List<Integer>> segments = new ArrayList<>();
            List<Integer> cur = new ArrayList<>();
            LocalDateTime prevEnd = null;
            for (int ii : ix) {
                TimelineEvent ev = events.get(ii);
                if (prevEnd != null && ev.start.isAfter(prevEnd)) {
                    segments.add(cur);
                    cur = new ArrayList<>();
                }
                cur.add(ii);
                prevEnd = ev.end;
            }
            if (!cur.isEmpty()) {
                segments.add(cur);
            }
            for (List<Integer> seg : segments) {
                if (seg.isEmpty()) {
                    continue;
                }
                out.put(barIdForIndices(seg), seg);
            }
        }
        return out;
    }

    static String barIdForIndices(List<Integer> indices) {
        if (indices == null || indices.isEmpty()) {
            return "";
        }
        int first = indices.getFirst();
        if (indices.size() == 1) {
            return "ev:" + first;
        }
        return "ev:" + first + "+" + (indices.size() - 1);
    }

    static boolean hasAssignableOperators(TimelineEvent e) {
        if (e == null) {
            return false;
        }
        if (!e.op.isBlank() || !e.sub.isBlank()) {
            return true;
        }
        return switch (e.eventKind) {
            case "machine_daily_startup",
                    "request_switch_prep",
                    "post_machining_cleanup",
                    "request_interval_buffer",
                    "break_resume_prep",
                    "machine_daily_inspection",
                    "daily_inspection" -> true;
            default -> false;
        };
    }

    static boolean subSplitStartup(String eventKind) {
        return switch (eventKind) {
            case "machine_daily_startup",
                    "request_switch_prep",
                    "post_machining_cleanup",
                    "request_interval_buffer",
                    "break_resume_prep",
                    "machine_daily_inspection",
                    "daily_inspection" -> true;
            default -> false;
        };
    }

    static List<EquipmentGanttAssignmentPerson> personsFromEvent(TimelineEvent e) {
        boolean startup = subSplitStartup(e.eventKind);
        List<String> raw =
                PersonNameBadgeText.orderedRawNamesFromOpSub(e.op, e.sub, startup);
        if (raw.isEmpty()) {
            return List.of();
        }
        List<EquipmentGanttAssignmentPerson> out = new ArrayList<>();
        for (int i = 0; i < raw.size(); i++) {
            EquipmentGanttAssignmentRole role =
                    i == 0 ? EquipmentGanttAssignmentRole.PRIMARY : EquipmentGanttAssignmentRole.SUB;
            out.add(EquipmentGanttAssignmentPerson.fromRawName(raw.get(i), role));
        }
        return List.copyOf(out);
    }

    private static List<EquipmentGanttAssignmentSlotBinding> buildSlotBindings(
            JsonTableIo.SheetTable table,
            List<List<String>> badgeSlotRows,
            List<LocalTime> slotStarts,
            List<TimelineEvent> events,
            Map<String, EquipmentGanttAssignmentBarUnit> barById) {
        if (table == null
                || badgeSlotRows == null
                || slotStarts == null
                || slotStarts.isEmpty()) {
            return List.of();
        }
        List<Map<String, String>> rows = table.rows();
        List<EquipmentGanttAssignmentSlotBinding> out = new ArrayList<>();
        for (int rowIdx = 0; rowIdx < rows.size(); rowIdx++) {
            Map<String, String> row = rows.get(rowIdx);
            if (row == null || isSectionBannerRow(row)) {
                continue;
            }
            if (rowIdx >= badgeSlotRows.size()) {
                continue;
            }
            List<String> timelineSlots =
                    EquipmentGanttTimelineColumns.timelineSlotTexts(row, table.columns());
            List<BarRunSlice> runs = EquipmentGanttPersonBadgeRunMerge.collectRuns(timelineSlots);
            LocalDate day = resolveRowDate(rows, rowIdx);
            String machine = row.getOrDefault(COL_MACH, "").strip();
            for (BarRunSlice run : runs) {
                String barId =
                        resolveBarIdForRun(day, machine, run, events, barById, slotStarts);
                if (barId.isBlank()) {
                    continue;
                }
                out.add(new EquipmentGanttAssignmentSlotBinding(rowIdx, run.fromSlot, run.toSlot, barId));
            }
        }
        return List.copyOf(out);
    }

    private static boolean isSectionBannerRow(Map<String, String> row) {
        String date = row.getOrDefault(COL_DATE, "").strip();
        String mach = row.getOrDefault(COL_MACH, "").strip();
        return !date.isEmpty() && mach.isEmpty();
    }

    private static LocalDate resolveRowDate(List<Map<String, String>> rows, int rowIdx) {
        for (int i = rowIdx; i >= 0; i--) {
            String d = rows.get(i).getOrDefault(COL_DATE, "").strip();
            if (!d.isEmpty()) {
                LocalDate parsed = parseSectionDate(d);
                if (parsed != null) {
                    return parsed;
                }
            }
        }
        return null;
    }

    private static LocalDate parseSectionDate(String banner) {
        String t = banner.strip();
        if (t.startsWith("【") && t.endsWith("】") && t.length() >= 2) {
            t = t.substring(1, t.length() - 1).strip();
        }
        if (t.length() >= 10 && t.charAt(4) == '/' && t.charAt(7) == '/') {
            try {
                return LocalDate.parse(
                        t.substring(0, 10).replace('/', '-'));
            } catch (RuntimeException ignored) {
                return null;
            }
        }
        return null;
    }

    private static String resolveBarIdForRun(
            LocalDate day,
            String machineDisplay,
            BarRunSlice run,
            List<TimelineEvent> events,
            Map<String, EquipmentGanttAssignmentBarUnit> barById,
            List<LocalTime> slotStarts) {
        if (day == null || run.text.isBlank()) {
            return "";
        }
        String taskKey = EquipmentGanttPersonBadgeRunMerge.mergeKey(run.text);
        LocalDateTime winStart = LocalDateTime.of(day, slotStarts.get(run.fromSlot));
        LocalDateTime winEnd =
                LocalDateTime.of(day, slotStarts.get(run.toSlot)).plusMinutes(10);
        for (EquipmentGanttAssignmentBarUnit unit : barById.values()) {
            if (unit.date() == null || !unit.date().equals(day)) {
                continue;
            }
            if (!machineDisplayMatches(unit.machine(), machineDisplay)) {
                continue;
            }
            if (!taskKeyMatchesUnit(taskKey, unit)) {
                continue;
            }
            for (int evIdx : unit.timelineEventIndices()) {
                TimelineEvent e = events.get(evIdx);
                if (rangesOverlap(e.start, e.end, winStart, winEnd)) {
                    return unit.barId();
                }
            }
        }
        return "";
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
                && EquipmentGanttAssignmentMetadata.slotKindMatchesEventKind(
                        kindPrefix, unit.eventKind())) {
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

    private static boolean rangesOverlap(
            LocalDateTime a0, LocalDateTime a1, LocalDateTime b0, LocalDateTime b1) {
        return a0.isBefore(b1) && b0.isBefore(a1);
    }

    static List<BarRunSlice> collectBarRunsForAssignment(List<String> slotTexts) {
        return EquipmentGanttPersonBadgeRunMerge.collectRuns(slotTexts);
    }

    @Deprecated
    static String personBadgeRunMergeKey(String slotText) {
        return EquipmentGanttPersonBadgeRunMerge.mergeKey(slotText);
    }

    record BarRunSlice(int fromSlot, int toSlot, String text) {}
}
