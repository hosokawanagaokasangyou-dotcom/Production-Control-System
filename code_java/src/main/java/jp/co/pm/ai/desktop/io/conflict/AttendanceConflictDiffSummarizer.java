package jp.co.pm.ai.desktop.io.conflict;

import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

/** attendance-data.json / 関連 Excel の業務語競合要約。 */
public final class AttendanceConflictDiffSummarizer implements ConflictDiffSummarizer {

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
                byte[] baseBytes = baselineSnapshots.getOrDefault(p, new byte[0]);
                byte[] diskB = diskBytes.getOrDefault(p, new byte[0]);
                if (baseBytes.length == 0 && diskB.length == 0) {
                    continue;
                }
                if (baseBytes.length == 0) {
                    lines.add("・" + name + " が新規に作成されています");
                    continue;
                }
                if (diskB.length == 0) {
                    lines.add("・" + name + " がディスク上にありません");
                    continue;
                }
                JsonNode base = JSON.readTree(new String(baseBytes, StandardCharsets.UTF_8));
                JsonNode disk = JSON.readTree(new String(diskB, StandardCharsets.UTF_8));
                lines.addAll(diffRoster(base, disk));
                lines.addAll(diffMemberCells(base, disk));
                lines.addAll(diffCompanyDays(base, disk));
            }
            if (lines.isEmpty()) {
                lines.add("・内容が変更されています（詳細キーは特定できませんでした）");
            }
            return String.join("\n", lines);
        } catch (Exception e) {
            return fallback(mismatched, diskBytes);
        }
    }

    private static String fallback(List<Path> mismatched, Map<Path, byte[]> diskBytes) {
        StringBuilder sb = new StringBuilder("詳細差分を生成できませんでした。\n");
        for (Path p : mismatched) {
            byte[] bytes = diskBytes.getOrDefault(p, new byte[0]);
            String hex = FileContentFingerprint.sha256Hex(bytes);
            String shortHex = hex.length() <= 12 ? hex : hex.substring(0, 12);
            sb.append("・")
                    .append(p.getFileName())
                    .append(" hash=")
                    .append(shortHex)
                    .append("…\n");
        }
        return sb.toString().trim();
    }

    private static Set<String> rosterNames(JsonNode root) {
        Set<String> names = new LinkedHashSet<>();
        JsonNode roster = root.path("member_roster");
        if (roster.isArray()) {
            for (JsonNode m : roster) {
                String n =
                        m.isTextual()
                                ? m.asText("").trim()
                                : m.path("name").asText("").trim();
                if (!n.isEmpty()) {
                    names.add(n);
                }
            }
        }
        return names;
    }

    private static List<String> diffRoster(JsonNode base, JsonNode disk) {
        Set<String> b = rosterNames(base);
        Set<String> d = rosterNames(disk);
        List<String> lines = new ArrayList<>();
        for (String n : d) {
            if (!b.contains(n)) {
                lines.add("・メンバー「" + n + "」が追加されています");
            }
        }
        for (String n : b) {
            if (!d.contains(n)) {
                lines.add("・メンバー「" + n + "」が削除されています");
            }
        }
        return lines;
    }

    /** member_attendance: { yyyy-MM-dd: { memberName: entry } } */
    private static List<String> diffMemberCells(JsonNode base, JsonNode disk) {
        int changed = 0;
        JsonNode b = base.path("member_attendance");
        JsonNode d = disk.path("member_attendance");
        if (d.isObject()) {
            Iterator<Map.Entry<String, JsonNode>> days = d.fields();
            while (days.hasNext()) {
                Map.Entry<String, JsonNode> day = days.next();
                JsonNode members = day.getValue();
                if (!members.isObject()) {
                    continue;
                }
                Iterator<Map.Entry<String, JsonNode>> memIt = members.fields();
                while (memIt.hasNext()) {
                    Map.Entry<String, JsonNode> mem = memIt.next();
                    JsonNode was = b.path(day.getKey()).path(mem.getKey());
                    if (!was.equals(mem.getValue())) {
                        changed++;
                    }
                }
            }
        }
        if (b.isObject()) {
            Iterator<Map.Entry<String, JsonNode>> days = b.fields();
            while (days.hasNext()) {
                Map.Entry<String, JsonNode> day = days.next();
                JsonNode members = day.getValue();
                if (!members.isObject()) {
                    continue;
                }
                Iterator<Map.Entry<String, JsonNode>> memIt = members.fields();
                while (memIt.hasNext()) {
                    Map.Entry<String, JsonNode> mem = memIt.next();
                    if (!d.path(day.getKey()).has(mem.getKey())) {
                        changed++;
                    }
                }
            }
        }
        if (changed == 0) {
            return List.of();
        }
        return List.of("・メンバー勤怠セルが " + changed + " 件変更されています");
    }

    private static List<String> diffCompanyDays(JsonNode base, JsonNode disk) {
        JsonNode b = base.path("company_calendar").path("days");
        JsonNode d = disk.path("company_calendar").path("days");
        if (!d.isObject() && !b.isObject()) {
            return List.of();
        }
        int changed = 0;
        Set<String> keys = new LinkedHashSet<>();
        if (b.isObject()) {
            b.fieldNames().forEachRemaining(keys::add);
        }
        if (d.isObject()) {
            d.fieldNames().forEachRemaining(keys::add);
        }
        for (String k : keys) {
            if (!b.path(k).equals(d.path(k))) {
                changed++;
            }
        }
        if (changed == 0) {
            return List.of();
        }
        return List.of("・会社カレンダーの日付区分が " + changed + " 件変更されています");
    }
}
