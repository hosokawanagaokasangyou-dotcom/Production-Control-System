package jp.co.pm.ai.desktop.dispatch.rules.history;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.time.Instant;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import jp.co.pm.ai.desktop.dispatch.rules.paths.DispatchRulePaths;

/** User-facing edit history snapshots. */
public final class DispatchRuleHistoryStore {

    private static final ObjectMapper JSON = new ObjectMapper();

    public record HistoryEntry(
            String id, String kind, String label, String savedAt, String summary, String snapshotFile) {}

    private DispatchRuleHistoryStore() {}

    public static Path indexPath(Map<String, String> ui) {
        return DispatchRulePaths.historyDirectory(ui).resolve("index.json");
    }

    public static HistoryEntry appendAutoSave(Map<String, String> ui, Path workJson) throws IOException {
        return appendSnapshot(ui, workJson, "auto_save", "保存");
    }

    public static HistoryEntry appendSnapshot(
            Map<String, String> ui, Path workJson, String kind, String label) throws IOException {
        Path hist = DispatchRulePaths.historyDirectory(ui);
        Path snaps = hist.resolve("snapshots");
        Files.createDirectories(snaps);
        String id = Instant.now().toString().replace(":", "").replace(".", "").substring(0, 15);
        String fileName = id + "_" + kind + ".json";
        Path snap = snaps.resolve(fileName);
        Files.copy(workJson, snap, StandardCopyOption.REPLACE_EXISTING);
        ObjectNode entry = JSON.createObjectNode();
        entry.put("id", id);
        entry.put("kind", kind);
        entry.put("label", label);
        entry.put("savedAt", Instant.now().toString());
        entry.put("schemaVersion", 1);
        entry.put("snapshotFile", "snapshots/" + fileName);
        entry.put("summary", label);
        ObjectNode index = readIndex(ui);
        ArrayNode entries = (ArrayNode) index.withArray("entries");
        entries.insert(0, entry);
        int max = index.path("maxEntries").asInt(50);
        while (entries.size() > max) {
            JsonNode old = entries.remove(entries.size() - 1);
            Path oldFile = hist.resolve(old.path("snapshotFile").asText(""));
            Files.deleteIfExists(oldFile);
        }
        Files.writeString(indexPath(ui), JSON.writerWithDefaultPrettyPrinter().writeValueAsString(index), StandardCharsets.UTF_8);
        return new HistoryEntry(id, kind, label, entry.path("savedAt").asText(), label, entry.path("snapshotFile").asText());
    }

    public static List<HistoryEntry> listEntries(Map<String, String> ui) throws IOException {
        ObjectNode index = readIndex(ui);
        List<HistoryEntry> out = new ArrayList<>();
        for (JsonNode n : index.withArray("entries")) {
            out.add(
                    new HistoryEntry(
                            n.path("id").asText(),
                            n.path("kind").asText(),
                            n.path("label").asText(),
                            n.path("savedAt").asText(),
                            n.path("summary").asText(),
                            n.path("snapshotFile").asText()));
        }
        return out;
    }

    public static void restore(Map<String, String> ui, Path workJson, String entryId) throws IOException {
        Path hist = DispatchRulePaths.historyDirectory(ui);
        ObjectNode index = readIndex(ui);
        JsonNode found = null;
        for (JsonNode n : index.withArray("entries")) {
            if (entryId.equals(n.path("id").asText())) {
                found = n;
                break;
            }
        }
        if (found == null) {
            throw new IOException("history entry not found: " + entryId);
        }
        appendSnapshot(ui, workJson, "auto_restore_guard", "復元前の自動退避");
        Path snap = hist.resolve(found.path("snapshotFile").asText());
        Files.copy(snap, workJson, StandardCopyOption.REPLACE_EXISTING);
    }

    private static ObjectNode readIndex(Map<String, String> ui) throws IOException {
        Path p = indexPath(ui);
        if (Files.isRegularFile(p)) {
            return (ObjectNode) JSON.readTree(Files.readString(p, StandardCharsets.UTF_8));
        }
        ObjectNode index = JSON.createObjectNode();
        index.put("version", 1);
        index.put("maxEntries", 50);
        index.set("entries", JSON.createArrayNode());
        return index;
    }
}
