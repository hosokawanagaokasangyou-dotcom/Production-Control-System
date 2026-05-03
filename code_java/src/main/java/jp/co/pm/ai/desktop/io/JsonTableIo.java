package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * Loads planning JSON exports (multi-sheet workbook or flat dispatch table) into plain tables.
 */
public final class JsonTableIo {

    private static final ObjectMapper JSON = new ObjectMapper();

    private JsonTableIo() {}

    public record SheetTable(List<String> columns, List<Map<String, String>> rows) {}

    /**
     * Workbook JSON with top-level {@code sheets} object ({@code member_schedule_*.json},
     * {@code production_plan_multi_day_*.json}).
     */
    public static Map<String, SheetTable> loadSheetsWorkbook(Path path) throws IOException {
        JsonNode root = JSON.readTree(Files.readString(path, StandardCharsets.UTF_8));
        JsonNode sheetsNode = root.get("sheets");
        if (sheetsNode == null || !sheetsNode.isObject()) {
            throw new IOException("JSON: sheets object missing");
        }
        Map<String, SheetTable> out = new LinkedHashMap<>();
        Iterator<Map.Entry<String, JsonNode>> it = sheetsNode.fields();
        while (it.hasNext()) {
            Map.Entry<String, JsonNode> en = it.next();
            SheetTable m = parseSheetTable(en.getValue());
            if (m != null) {
                out.put(en.getKey(), m);
            }
        }
        return out;
    }

    /** Flat single-table JSON (result dispatch export; columns + rows array). */
    public static SheetTable loadFlatTable(Path path) throws IOException {
        JsonNode root = JSON.readTree(Files.readString(path, StandardCharsets.UTF_8));
        JsonNode columnsNode = root.get("columns");
        JsonNode rowsNode = root.get("rows");
        if (columnsNode == null
                || !columnsNode.isArray()
                || rowsNode == null
                || !rowsNode.isArray()) {
            throw new IOException("JSON: columns/rows missing");
        }
        List<String> headerOrder = new ArrayList<>();
        for (JsonNode c : columnsNode) {
            headerOrder.add(c != null && c.isTextual() ? c.asText("") : "");
        }
        List<Map<String, String>> rowMaps = new ArrayList<>();
        for (JsonNode r : rowsNode) {
            if (r == null || !r.isObject()) {
                continue;
            }
            LinkedHashMap<String, String> row = new LinkedHashMap<>();
            for (String h : headerOrder) {
                row.put(h, formatCell(r.get(h)));
            }
            rowMaps.add(row);
        }
        return new SheetTable(List.copyOf(headerOrder), List.copyOf(rowMaps));
    }

    static SheetTable parseSheetTable(JsonNode sheetNode) {
        if (sheetNode == null || !sheetNode.isObject()) {
            return null;
        }
        JsonNode columnsNode = sheetNode.get("columns");
        JsonNode rowsNode = sheetNode.get("rows");
        if (columnsNode == null
                || !columnsNode.isArray()
                || rowsNode == null
                || !rowsNode.isArray()) {
            return null;
        }
        List<String> columns = new ArrayList<>();
        for (JsonNode c : columnsNode) {
            columns.add(c != null && c.isTextual() ? c.asText("") : "");
        }
        List<Map<String, String>> rowMaps = new ArrayList<>();
        for (JsonNode r : rowsNode) {
            if (r == null || !r.isObject()) {
                continue;
            }
            LinkedHashMap<String, String> row = new LinkedHashMap<>();
            for (String h : columns) {
                row.put(h, formatCell(r.get(h)));
            }
            rowMaps.add(row);
        }
        return new SheetTable(List.copyOf(columns), List.copyOf(rowMaps));
    }

    /** Same rules as UI spreadsheet viewers (ISO date strings shortened to yyyy-MM-dd). */
    public static String formatCell(JsonNode n) {
        if (n == null || n.isNull()) {
            return "";
        }
        if (n.isBoolean()) {
            return n.asBoolean() ? "true" : "false";
        }
        if (n.isInt() || n.isLong()) {
            return Long.toString(n.longValue());
        }
        if (n.isDouble() || n.isFloat() || n.isBigDecimal()) {
            double d = n.asDouble();
            if (Double.isFinite(d) && d == Math.rint(d) && Math.abs(d) < 1e15) {
                return Long.toString((long) d);
            }
            return n.asText("");
        }
        if (n.isTextual()) {
            String t = n.asText("");
            if (t.length() >= 19 && t.charAt(10) == 'T' && t.charAt(4) == '-') {
                return t.substring(0, 10);
            }
            return t;
        }
        return n.asText("");
    }

    /**
     * Distinct operator names from workbook sheet keys (excluding sheets without the time-slot column).
     */
    public static List<String> memberOperatorNames(Map<String, SheetTable> memberSheets) {
        final String timeCol = "\u6642\u9593\u5e2f";
        Set<String> keepOrder = new LinkedHashSet<>();
        for (Map.Entry<String, SheetTable> e : memberSheets.entrySet()) {
            if (e.getValue().columns().contains(timeCol)) {
                keepOrder.add(e.getKey());
            }
        }
        return List.copyOf(keepOrder);
    }
}
