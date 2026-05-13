package jp.co.pm.ai.planning.stage2.output;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

import jp.co.pm.ai.desktop.io.ExcelCellReadSupport;

/**
 * Python {@code workbook_payload_from_final_xlsx_file} / {@code plan_workbook_sidecar} のミラー JSON に近い形で
 * {@code format_version:2} ペイロードを書き出す。
 */
public final class Stage2WorkbookJsonWriter {

    public static final int WORKBOOK_JSON_FORMAT_VERSION = 2;

    private static final ObjectMapper MAPPER = new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);

    private Stage2WorkbookJsonWriter() {}

    public static void writeFromXlsx(Path xlsxPath, Path jsonPath, Map<String, Object> metadataExtra)
            throws IOException {
        Map<String, Object> payload = buildPayloadFromXlsx(xlsxPath, metadataExtra);
        Files.createDirectories(jsonPath.getParent());
        try (OutputStream os = Files.newOutputStream(jsonPath)) {
            MAPPER.writeValue(os, payload);
        }
    }

    public static void writePayload(Path jsonPath, Map<String, Object> payload) throws IOException {
        Files.createDirectories(jsonPath.getParent());
        try (OutputStream os = Files.newOutputStream(jsonPath)) {
            MAPPER.writeValue(os, payload);
        }
    }

    public static Map<String, Object> buildPayloadFromXlsx(Path xlsxPath, Map<String, Object> metadataExtra)
            throws IOException {
        String base = xlsxPath.getFileName() != null ? xlsxPath.getFileName().toString() : xlsxPath.toString();
        DataFormatter fmt = new DataFormatter(Locale.ROOT);
        Map<String, Object> sheets = new LinkedHashMap<>();
        try (Workbook wb = WorkbookFactory.create(xlsxPath.toFile())) {
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                Sheet sh = wb.getSheetAt(i);
                String name = sh.getSheetName();
                sheets.put(name, sheetToTabular(sh, fmt));
            }
        }
        Map<String, Object> payload = new LinkedHashMap<>();
        payload.put("format_version", WORKBOOK_JSON_FORMAT_VERSION);
        payload.put("source_xlsx", base);
        payload.put("sheets", sheets);
        if (metadataExtra != null && !metadataExtra.isEmpty()) {
            payload.putAll(metadataExtra);
        }
        return payload;
    }

    private static Map<String, Object> sheetToTabular(Sheet sh, DataFormatter fmt) {
        Row h = sh.getRow(0);
        List<String> columns = new ArrayList<>();
        if (h != null) {
            short last = h.getLastCellNum();
            for (int c = 0; c < last; c++) {
                columns.add(cellStr(fmt, h.getCell(c)));
            }
            while (!columns.isEmpty() && columns.get(columns.size() - 1).isBlank()) {
                columns.remove(columns.size() - 1);
            }
        }
        List<Map<String, Object>> rows = new ArrayList<>();
        for (int r = 1; r <= sh.getLastRowNum(); r++) {
            Row row = sh.getRow(r);
            Map<String, Object> rec = new LinkedHashMap<>();
            for (int c = 0; c < columns.size(); c++) {
                String col = columns.get(c);
                String v =
                        row == null
                                ? ""
                                : ExcelCellReadSupport.normalizeCommaDigitArtifacts(
                                        fmt.formatCellValue(row.getCell(c)).strip());
                rec.put(col, v);
            }
            rows.add(rec);
        }
        Map<String, Object> out = new LinkedHashMap<>();
        out.put("columns", columns);
        out.put("row_count", rows.size());
        out.put("rows", rows);
        return out;
    }

    private static String cellStr(DataFormatter fmt, Cell cell) {
        if (cell == null) {
            return "";
        }
        return fmt.formatCellValue(cell).strip();
    }
}
