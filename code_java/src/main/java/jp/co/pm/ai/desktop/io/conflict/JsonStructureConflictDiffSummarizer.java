package jp.co.pm.ai.desktop.io.conflict;

import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

/** JSON の浅い構造差分を業務語で要約する。 */
public final class JsonStructureConflictDiffSummarizer implements ConflictDiffSummarizer {

    private static final ObjectMapper JSON = new ObjectMapper();

    private final String screenLabel;

    public JsonStructureConflictDiffSummarizer(String screenLabel) {
        this.screenLabel = screenLabel != null ? screenLabel : "JSON";
    }

    @Override
    public String summarize(
            Map<Path, byte[]> baselineSnapshots,
            Map<Path, byte[]> diskBytes,
            List<Path> mismatched) {
        List<String> lines = new ArrayList<>();
        lines.add(screenLabel + " の保存先が外部で変更されています。");
        try {
            for (Path p : mismatched) {
                String name = p.getFileName() != null ? p.getFileName().toString() : p.toString();
                if (name.endsWith(".xlsx") || name.endsWith(".xlsm") || name.endsWith(".xls")) {
                    lines.add("・関連ブック（" + name + "）が変更されています");
                    continue;
                }
                byte[] baseBytes = baselineSnapshots.getOrDefault(p, new byte[0]);
                byte[] diskB = diskBytes.getOrDefault(p, new byte[0]);
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
                lines.addAll(diffShallow(base, disk, name));
            }
            if (lines.size() == 1) {
                lines.add("・内容が変更されています");
            }
            return String.join("\n", lines);
        } catch (Exception e) {
            return screenLabel
                    + " の詳細差分を生成できませんでした。\n・"
                    + mismatched.stream()
                            .map(x -> x.getFileName() != null ? x.getFileName().toString() : x.toString())
                            .reduce((a, b) -> a + ", " + b)
                            .orElse("");
        }
    }

    private static List<String> diffShallow(JsonNode base, JsonNode disk, String fileName) {
        List<String> lines = new ArrayList<>();
        if (!base.isObject() || !disk.isObject()) {
            lines.add("・" + fileName + " の内容が変更されています");
            return lines;
        }
        Set<String> keys = new LinkedHashSet<>();
        base.fieldNames().forEachRemaining(keys::add);
        disk.fieldNames().forEachRemaining(keys::add);
        for (String k : keys) {
            boolean inBase = base.has(k);
            boolean inDisk = disk.has(k);
            if (!inBase && inDisk) {
                lines.add("・キー「" + k + "」が追加されています");
                continue;
            }
            if (inBase && !inDisk) {
                lines.add("・キー「" + k + "」が削除されています");
                continue;
            }
            JsonNode bv = base.get(k);
            JsonNode dv = disk.get(k);
            if (bv.isArray() && dv.isArray() && bv.size() != dv.size()) {
                lines.add("・配列「" + k + "」の要素数が " + bv.size() + " → " + dv.size());
            } else if (!bv.equals(dv)) {
                lines.add("・「" + k + "」の値が変更されています");
            }
        }
        return lines;
    }
}
