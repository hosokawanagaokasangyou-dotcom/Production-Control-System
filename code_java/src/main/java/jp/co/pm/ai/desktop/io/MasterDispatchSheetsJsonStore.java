package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

/** {@link MasterDispatchSheetsDocument} の UTF-8 JSON 読み書き。 */
public final class MasterDispatchSheetsJsonStore {

    private static final ObjectMapper JSON = new ObjectMapper();

    private MasterDispatchSheetsJsonStore() {}

    public static void write(Path path, MasterDispatchSheetsDocument doc) throws IOException {
        Objects.requireNonNull(path, "path");
        Objects.requireNonNull(doc, "doc");
        Path parent = path.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        ObjectNode root = JSON.createObjectNode();
        root.put("schemaVersion", doc.schemaVersion());
        root.put("factorySite", doc.factorySite());
        root.put("sourceWorkbook", doc.sourceWorkbook());
        root.put("importedAt", doc.importedAt());
        ObjectNode sheets = root.putObject("sheets");
        for (String key : MasterDispatchSheetsDocument.SHEET_KEYS) {
            MasterDispatchSheetsDocument.SheetGrid grid = doc.sheet(key);
            ObjectNode sn = sheets.putObject(key);
            sn.put("sheetName", grid.sheetName());
            ArrayNode rows = sn.putArray("rows");
            for (List<String> row : grid.rows()) {
                ArrayNode rr = rows.addArray();
                for (String c : row) {
                    rr.add(c != null ? c : "");
                }
            }
        }
        Files.writeString(
                path,
                JSON.writerWithDefaultPrettyPrinter().writeValueAsString(root) + "\n",
                StandardCharsets.UTF_8);
    }

    public static MasterDispatchSheetsDocument read(Path path) throws IOException {
        Objects.requireNonNull(path, "path");
        JsonNode root = JSON.readTree(Files.readString(path, StandardCharsets.UTF_8));
        int version = root.path("schemaVersion").asInt(MasterDispatchSheetsDocument.SCHEMA_VERSION);
        String factory = root.path("factorySite").asText("");
        String source = root.path("sourceWorkbook").asText("");
        String importedAt = root.path("importedAt").asText("");
        JsonNode sheetsNode = root.path("sheets");
        Map<String, MasterDispatchSheetsDocument.SheetGrid> sheets = new LinkedHashMap<>();
        if (sheetsNode != null && sheetsNode.isObject()) {
            Iterator<Map.Entry<String, JsonNode>> it = sheetsNode.fields();
            while (it.hasNext()) {
                Map.Entry<String, JsonNode> en = it.next();
                sheets.put(en.getKey(), parseGrid(en.getKey(), en.getValue()));
            }
        }
        return new MasterDispatchSheetsDocument(version, factory, source, importedAt, sheets);
    }

    private static MasterDispatchSheetsDocument.SheetGrid parseGrid(String key, JsonNode node) {
        String name =
                node != null
                        ? node.path("sheetName").asText(MasterDispatchSheetsDocument.defaultSheetName(key))
                        : MasterDispatchSheetsDocument.defaultSheetName(key);
        List<List<String>> rows = new ArrayList<>();
        JsonNode rowsNode = node != null ? node.get("rows") : null;
        if (rowsNode != null && rowsNode.isArray()) {
            for (JsonNode r : rowsNode) {
                List<String> row = new ArrayList<>();
                if (r != null && r.isArray()) {
                    for (JsonNode c : r) {
                        row.add(c != null && !c.isNull() ? c.asText("") : "");
                    }
                }
                rows.add(row);
            }
        }
        return new MasterDispatchSheetsDocument.SheetGrid(name, rows);
    }
}
