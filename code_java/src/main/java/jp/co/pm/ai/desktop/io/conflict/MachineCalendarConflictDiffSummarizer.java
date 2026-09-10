package jp.co.pm.ai.desktop.io.conflict;

import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

/** machine-calendar-data.json / 関連 Excel の業務語競合要約。 */
public final class MachineCalendarConflictDiffSummarizer implements ConflictDiffSummarizer {

    private static final ObjectMapper JSON = new ObjectMapper();

    @Override
    public String summarize(
            Map<Path, byte[]> baselineSnapshots,
            Map<Path, byte[]> diskBytes,
            List<Path> mismatched) {
        List<String> lines = new ArrayList<>();
        try {
            for (Path p : mismatched) {
                String name = p.getFileName() != null ? p.getFileName().toString() : p.toString();
                if (name.endsWith(".xlsx") || name.endsWith(".xlsm")) {
                    lines.add("・関連 Excel（" + name + "）が変更されています");
                    continue;
                }
                byte[] baseBytes = SnapshotPresence.get(baselineSnapshots, p);
                byte[] diskB = SnapshotPresence.get(diskBytes, p);
                if (SnapshotPresence.isAbsent(baseBytes)) {
                    lines.add("・" + name + " が新規に作成されています");
                    continue;
                }
                if (SnapshotPresence.isAbsent(diskB)) {
                    lines.add("・" + name + " がディスク上にありません");
                    continue;
                }
                JsonNode base = JSON.readTree(new String(baseBytes, StandardCharsets.UTF_8));
                JsonNode disk = JSON.readTree(new String(diskB, StandardCharsets.UTF_8));
                lines.addAll(diffColumns(base, disk));
                lines.addAll(diffOccupancy(base, disk));
            }
            if (lines.isEmpty()) {
                lines.add("・機械カレンダー内容が変更されています");
            }
            return String.join("\n", lines);
        } catch (Exception e) {
            return SaveConflictGate.fallbackSummary(diskBytes, ConflictCheckResult.conflict(mismatched));
        }
    }

    private static List<String> diffColumns(JsonNode base, JsonNode disk) {
        int b = base.path("columns").isArray() ? base.path("columns").size() : 0;
        int d = disk.path("columns").isArray() ? disk.path("columns").size() : 0;
        if (b == d) {
            return List.of();
        }
        return List.of("・機械列（columns）が " + b + " → " + d + " に変化しています");
    }

    private static List<String> diffOccupancy(JsonNode base, JsonNode disk) {
        int changed = countLeafDiff(base.path("occupancy"), disk.path("occupancy"));
        if (changed == 0) {
            return List.of();
        }
        return List.of("・稼働／占有スロットが " + changed + " 件変更されています");
    }

    private static int countLeafDiff(JsonNode a, JsonNode b) {
        if (a.equals(b)) {
            return 0;
        }
        if (!a.isObject() && !b.isObject()) {
            return 1;
        }
        int n = 0;
        java.util.Set<String> keys = new java.util.LinkedHashSet<>();
        if (a.isObject()) {
            a.fieldNames().forEachRemaining(keys::add);
        }
        if (b.isObject()) {
            b.fieldNames().forEachRemaining(keys::add);
        }
        for (String k : keys) {
            n += countLeafDiff(a.path(k), b.path(k));
        }
        return n;
    }
}
