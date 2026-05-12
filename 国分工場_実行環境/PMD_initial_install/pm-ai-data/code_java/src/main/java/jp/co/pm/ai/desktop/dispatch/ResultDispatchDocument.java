package jp.co.pm.ai.desktop.dispatch;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/** Canonical long-format document (same logical rows as result dispatch JSON). */
public final class ResultDispatchDocument {

    private int formatVersion = 1;
    private String sheetName = "結果_配台表";
    private String excelTableName = "_t結果_配台表";
    private final List<String> columns;
    private final List<Map<String, String>> rows;

    public ResultDispatchDocument(List<String> columns, List<Map<String, String>> rows) {
        this.columns = new ArrayList<>(columns);
        this.rows = new ArrayList<>(rows);
    }

    public static ResultDispatchDocument empty() {
        return new ResultDispatchDocument(ResultDispatchSchema.canonicalColumnOrder(), new ArrayList<>());
    }

    public int formatVersion() {
        return formatVersion;
    }

    public void setFormatVersion(int formatVersion) {
        this.formatVersion = formatVersion;
    }

    public String sheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName != null ? sheetName : "";
    }

    public String excelTableName() {
        return excelTableName;
    }

    public void setExcelTableName(String excelTableName) {
        this.excelTableName = excelTableName != null ? excelTableName : "";
    }

    public List<String> columns() {
        return columns;
    }

    public List<Map<String, String>> rows() {
        return rows;
    }

    /** Deep-enough copy for undo or branching (maps are new instances). */
    public ResultDispatchDocument copy() {
        List<Map<String, String>> copyRows = new ArrayList<>();
        for (Map<String, String> r : rows) {
            copyRows.add(new LinkedHashMap<>(r));
        }
        ResultDispatchDocument d = new ResultDispatchDocument(new ArrayList<>(columns), copyRows);
        d.formatVersion = formatVersion;
        d.sheetName = sheetName;
        d.excelTableName = excelTableName;
        return d;
    }
}
