package jp.co.pm.ai.desktop.dispatch;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * Equipment calendar blocks: days where the machine column has at least one occupied slot (planning_core).
 */
public final class MachineCalendarBlockIndex {

    private static final ObjectMapper JSON = new ObjectMapper();

    private final Map<String, Set<LocalDate>> blocksByEquipmentKey;

    public MachineCalendarBlockIndex(Map<String, Set<LocalDate>> blocksByEquipmentKey) {
        this.blocksByEquipmentKey = Map.copyOf(blocksByEquipmentKey);
    }

    public boolean isEmpty() {
        return blocksByEquipmentKey.isEmpty();
    }

    public static MachineCalendarBlockIndex empty() {
        return new MachineCalendarBlockIndex(Map.of());
    }

    /**
     * Runs {@code export_machine_calendar_blocks.py} from {@code pythonScriptDir}.
     *
     * @return empty index on failure (drops allowed everywhere calendar unknown).
     */
    public static MachineCalendarBlockIndex load(Path masterWorkbook, Path pythonExe, Path pythonScriptDir)
            throws Exception {
        if (masterWorkbook == null || !Files.isRegularFile(masterWorkbook)) {
            return empty();
        }
        Path script = pythonScriptDir.resolve("export_machine_calendar_blocks.py");
        if (!Files.isRegularFile(script)) {
            return empty();
        }
        ProcessBuilder pb =
                new ProcessBuilder(
                        pythonExe.toString(),
                        script.toString(),
                        masterWorkbook.toAbsolutePath().toString());
        pb.directory(pythonScriptDir.toFile());
        pb.redirectErrorStream(true);
        Process p = pb.start();
        String out;
        try (BufferedReader br =
                new BufferedReader(new InputStreamReader(p.getInputStream(), StandardCharsets.UTF_8))) {
            StringBuilder sb = new StringBuilder();
            String line;
            while ((line = br.readLine()) != null) {
                sb.append(line);
            }
            out = sb.toString();
        }
        boolean finished = p.waitFor(120, TimeUnit.SECONDS);
        if (!finished) {
            p.destroyForcibly();
            return empty();
        }
        if (p.exitValue() != 0) {
            return empty();
        }
        return parseStdoutJson(out);
    }

    static MachineCalendarBlockIndex parseStdoutJson(String stdout) {
        try {
            JsonNode root = JSON.readTree(stdout);
            JsonNode blocks = root.get("blocks");
            if (!(blocks instanceof ObjectNode on)) {
                return empty();
            }
            Map<String, Set<LocalDate>> acc = new HashMap<>();
            on.fieldNames()
                    .forEachRemaining(
                            key -> {
                                Set<LocalDate> ds = new HashSet<>();
                                JsonNode arr = on.get(key);
                                if (arr != null && arr.isArray()) {
                                    for (JsonNode n : arr) {
                                        if (n.isTextual()) {
                                            String t = n.asText("");
                                            if (t.length() >= 10) {
                                                LocalDate d = LocalDate.parse(t.substring(0, 10));
                                                ds.add(d);
                                            }
                                        }
                                    }
                                }
                                if (!ds.isEmpty()) {
                                    acc.put(key.trim(), ds);
                                }
                            });
            return new MachineCalendarBlockIndex(acc);
        } catch (Exception e) {
            return empty();
        }
    }

    /**
     * Whether dropping allocation quantity onto this process/machine/day should be denied (machine occupied /
     * blocked day).
     */
    public boolean isBlockedDay(String processName, String machineName, LocalDate day) {
        if (blocksByEquipmentKey.isEmpty()) {
            return false;
        }
        String p = nz(processName).trim();
        String m = nz(machineName).trim();
        for (Map.Entry<String, Set<LocalDate>> e : blocksByEquipmentKey.entrySet()) {
            if (!e.getValue().contains(day)) {
                continue;
            }
            if (matchesEquipmentKey(e.getKey(), p, m)) {
                return true;
            }
        }
        return false;
    }

    static boolean matchesEquipmentKey(String equipmentKey, String process, String machine) {
        String k = equipmentKey.trim();
        if (k.isEmpty()) {
            return false;
        }
        if (k.equals(machine)) {
            return true;
        }
        if (k.equals(process + "\t" + machine)) {
            return true;
        }
        String compactK = k.replaceAll("\\s+", "");
        String compactPM = (process + machine).replaceAll("\\s+", "");
        if (!compactPM.isEmpty() && compactK.equals(compactPM)) {
            return true;
        }
        if (!machine.isEmpty() && k.contains(machine.trim())) {
            if (process.isEmpty() || k.contains(process.trim())) {
                return true;
            }
        }
        return false;
    }

    private static String nz(String s) {
        return s != null ? s : "";
    }
}
