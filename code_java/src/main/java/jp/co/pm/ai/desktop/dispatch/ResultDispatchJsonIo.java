package jp.co.pm.ai.desktop.dispatch;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * Read / write {@code ????_?z??\.json} compatible with {@link jp.co.pm.ai.desktop.ResultDispatchTableTabController}.
 */
public final class ResultDispatchJsonIo {

    private static final ObjectMapper JSON = new ObjectMapper();

    private ResultDispatchJsonIo() {}

    public static ResultDispatchDocument read(Path path) throws Exception {
        String raw = Files.readString(path, StandardCharsets.UTF_8);
        JsonNode root = JSON.readTree(raw);
        int formatVer = root.path("format_version").asInt(1);
        String sheetName = textOr(root, "sheet_name");
        String excelTable = textOr(root, "excel_table_name");
        JsonNode columnsNode = root.get("columns");
        JsonNode rowsNode = root.get("rows");
        if (columnsNode == null || !columnsNode.isArray()) {
            throw new IllegalArgumentException("missing columns array");
        }
        List<String> headerOrder = new ArrayList<>();
        for (JsonNode c : columnsNode) {
            if (c != null && c.isTextual()) {
                headerOrder.add(c.asText(""));
            }
        }
        List<Map<String, String>> rowMaps = new ArrayList<>();
        if (rowsNode != null && rowsNode.isArray()) {
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
        }
        ResultDispatchDocument doc = new ResultDispatchDocument(headerOrder, rowMaps);
        doc.setFormatVersion(formatVer);
        doc.setSheetName(sheetName);
        doc.setExcelTableName(excelTable);
        ResultDispatchNormalizer.normalizeInPlace(doc.columns(), doc.rows());
        return doc;
    }

    public static void write(Path path, ResultDispatchDocument doc) throws Exception {
        ObjectNode root = JSON.createObjectNode();
        root.put("format_version", doc.formatVersion());
        root.put("sheet_name", doc.sheetName());
        root.put("excel_table_name", doc.excelTableName());
        ArrayNode cols = JSON.createArrayNode();
        for (String c : doc.columns()) {
            cols.add(c);
        }
        root.set("columns", cols);
        root.put("row_count", doc.rows().size());
        ArrayNode rows = JSON.createArrayNode();
        for (Map<String, String> row : doc.rows()) {
            ObjectNode o = JSON.createObjectNode();
            for (String h : doc.columns()) {
                putCell(o, h, row.get(h));
            }
            rows.add(o);
        }
        root.set("rows", rows);
        String text = JSON.writerWithDefaultPrettyPrinter().writeValueAsString(root) + "\n";
        Files.createDirectories(path.getParent());
        Files.writeString(path, text, StandardCharsets.UTF_8);
    }

    private static void putCell(ObjectNode o, String key, String val) {
        if (val == null || val.isBlank()) {
            o.putNull(key);
            return;
        }
        String t = val.trim();
        if (ResultDispatchSchema.isDateColumn(key)) {
            o.put(key, t);
            return;
        }
        if (ResultDispatchSchema.COL_DISPATCH_QTY.equals(key)) {
            try {
                if (t.contains(".") || t.contains("e") || t.contains("E")) {
                    o.put(key, Double.parseDouble(t.replace(",", "")));
                } else {
                    o.put(key, Long.parseLong(t.replace(",", "")));
                }
            } catch (NumberFormatException e) {
                o.put(key, t);
            }
            return;
        }
        try {
            if (t.contains(".") || t.contains("e") || t.contains("E")) {
                o.put(key, Double.parseDouble(t.replace(",", "")));
            } else {
                o.put(key, Long.parseLong(t.replace(",", "")));
            }
        } catch (NumberFormatException e) {
            o.put(key, val);
        }
    }

    private static String textOr(JsonNode n, String field) {
        JsonNode x = n.get(field);
        return x == null || x.isNull() ? "" : x.asText("");
    }

    private static String formatCell(JsonNode n) {
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
}
