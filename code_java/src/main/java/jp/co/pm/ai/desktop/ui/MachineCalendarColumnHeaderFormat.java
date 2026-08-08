package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/** need シート由来の工程×機械列見出しをグリッド向けに短く整形する。 */
final class MachineCalendarColumnHeaderFormat {

    private static final List<String> KNOWN_FACTORY_SUFFIXES =
            List.of("　湖南", " 湖南", "湖南");

    record Display(String text, String tooltip) {}

    private MachineCalendarColumnHeaderFormat() {}

    static List<Display> formatAll(List<EditableMachineCalendarGridPane.ColumnDef> columns) {
        if (columns == null || columns.isEmpty()) {
            return List.of();
        }
        String commonSuffix = detectCommonMachineSuffix(columns);
        List<String> rawLabels = new ArrayList<>();
        List<String> tooltips = new ArrayList<>();
        for (EditableMachineCalendarGridPane.ColumnDef col : columns) {
            String process = normalize(col.process());
            String machine = normalize(col.machine());
            String machineShort = stripSuffix(machine, commonSuffix);
            if (machineShort.isBlank() && !machine.isBlank()) {
                machineShort = machine;
            }
            String tooltip = buildTooltip(process, machine, col.equipmentKey());
            tooltips.add(tooltip);
            if (isProcessRedundant(process, machineShort)) {
                rawLabels.add(firstNonBlank(machineShort, process, col.equipmentKey()));
            } else if (process.isBlank()) {
                rawLabels.add(firstNonBlank(machineShort, col.equipmentKey()));
            } else if (machineShort.isBlank()) {
                rawLabels.add(process);
            } else {
                rawLabels.add(process + "·" + machineShort);
            }
        }
        Map<String, Integer> counts = new HashMap<>();
        for (String label : rawLabels) {
            counts.merge(label, 1, Integer::sum);
        }
        List<Display> out = new ArrayList<>();
        for (int i = 0; i < columns.size(); i++) {
            EditableMachineCalendarGridPane.ColumnDef col = columns.get(i);
            String process = normalize(col.process());
            String machine = normalize(col.machine());
            String machineShort = stripSuffix(machine, commonSuffix);
            if (machineShort.isBlank() && !machine.isBlank()) {
                machineShort = machine;
            }
            String label = rawLabels.get(i);
            if (counts.getOrDefault(label, 0) > 1) {
                label = disambiguateLabel(process, machineShort, col.equipmentKey());
            }
            out.add(new Display(label, tooltips.get(i)));
        }
        return out;
    }

    private static String disambiguateLabel(
            String process, String machineShort, String equipmentKey) {
        if (!process.isBlank() && !machineShort.isBlank()) {
            return process + "·" + machineShort;
        }
        return firstNonBlank(process, machineShort, equipmentKey);
    }

    private static String buildTooltip(String process, String machine, String equipmentKey) {
        if (!process.isBlank() && !machine.isBlank()) {
            return process + " / " + machine + "\n" + equipmentKey;
        }
        if (!machine.isBlank()) {
            return machine + "\n" + equipmentKey;
        }
        if (!process.isBlank()) {
            return process + "\n" + equipmentKey;
        }
        return equipmentKey;
    }

    private static String detectCommonMachineSuffix(
            List<EditableMachineCalendarGridPane.ColumnDef> columns) {
        List<String> machines = new ArrayList<>();
        for (EditableMachineCalendarGridPane.ColumnDef col : columns) {
            String m = normalize(col.machine());
            if (!m.isBlank()) {
                machines.add(m);
            }
        }
        if (machines.size() < 2) {
            return "";
        }
        for (String suffix : KNOWN_FACTORY_SUFFIXES) {
            boolean allMatch = true;
            for (String m : machines) {
                if (!m.endsWith(suffix)) {
                    allMatch = false;
                    break;
                }
            }
            if (allMatch) {
                return suffix;
            }
        }
        return "";
    }

    private static String stripSuffix(String machine, String suffix) {
        if (machine.isBlank() || suffix.isBlank()) {
            return machine;
        }
        if (machine.endsWith(suffix)) {
            return machine.substring(0, machine.length() - suffix.length()).trim();
        }
        return machine;
    }

    private static boolean isProcessRedundant(String process, String machineShort) {
        if (process.isBlank() || machineShort.isBlank()) {
            return true;
        }
        if (process.equals(machineShort)) {
            return true;
        }
        if (machineShort.startsWith(process)) {
            return true;
        }
        if (machineShort.contains(process)) {
            return true;
        }
        String machineNoMachineChar = machineShort.replace("機", "");
        if (!machineNoMachineChar.isBlank() && machineNoMachineChar.equals(process)) {
            return true;
        }
        return false;
    }

    private static String normalize(String s) {
        return s == null ? "" : s.strip();
    }

    private static String firstNonBlank(String... values) {
        for (String v : values) {
            if (v != null && !v.isBlank()) {
                return v;
            }
        }
        return "";
    }
}
