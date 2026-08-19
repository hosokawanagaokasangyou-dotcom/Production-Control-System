package jp.co.pm.ai.desktop.io;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * master.xlsm の skills / need / speed / 組み合わせ表を格子 JSON にした文書。
 */
public record MasterDispatchSheetsDocument(
        int schemaVersion,
        String factorySite,
        String sourceWorkbook,
        String importedAt,
        Map<String, SheetGrid> sheets) {

    public static final int SCHEMA_VERSION = 1;
    public static final String KEY_SKILLS = "skills";
    public static final String KEY_NEED = "need";
    public static final String KEY_SPEED = "speed";
    public static final String KEY_TEAM_COMBINATIONS = "teamCombinations";

    public static final List<String> SHEET_KEYS =
            List.of(KEY_SKILLS, KEY_NEED, KEY_SPEED, KEY_TEAM_COMBINATIONS);

    public MasterDispatchSheetsDocument {
        factorySite = factorySite != null ? factorySite : "";
        sourceWorkbook = sourceWorkbook != null ? sourceWorkbook : "";
        importedAt = importedAt != null ? importedAt : "";
        LinkedHashMap<String, SheetGrid> copy = new LinkedHashMap<>();
        Map<String, SheetGrid> src = sheets != null ? sheets : Map.of();
        for (String key : SHEET_KEYS) {
            SheetGrid g = src.get(key);
            copy.put(key, g != null ? g : SheetGrid.empty(defaultSheetName(key)));
        }
        sheets = Map.copyOf(copy);
    }

    public SheetGrid sheet(String key) {
        SheetGrid g = sheets.get(key);
        return g != null ? g : SheetGrid.empty(defaultSheetName(key));
    }

    public static String defaultSheetName(String key) {
        if (KEY_TEAM_COMBINATIONS.equals(key)) {
            return "組み合わせ表";
        }
        if (KEY_NEED.equals(key)) {
            return "need";
        }
        if (KEY_SPEED.equals(key)) {
            return "speed";
        }
        return "skills";
    }

    public static MasterDispatchSheetsDocument empty(String factorySite) {
        return new MasterDispatchSheetsDocument(
                SCHEMA_VERSION, factorySite, "", "", Map.of());
    }

    /** 末尾の空行・空列を落とす（保存時・表示抽出用）。 */
    public static List<List<String>> trimTrailingEmpty(List<List<String>> raw) {
        if (raw == null || raw.isEmpty()) {
            return List.of();
        }
        int lastUsedRow = -1;
        int lastUsedCol = -1;
        for (int r = 0; r < raw.size(); r++) {
            List<String> row = raw.get(r);
            if (row == null) {
                continue;
            }
            for (int c = 0; c < row.size(); c++) {
                String v = row.get(c);
                if (v != null && !v.isEmpty()) {
                    lastUsedRow = r;
                    if (c > lastUsedCol) {
                        lastUsedCol = c;
                    }
                }
            }
        }
        if (lastUsedRow < 0) {
            return List.of();
        }
        int width = lastUsedCol + 1;
        List<List<String>> out = new ArrayList<>(lastUsedRow + 1);
        for (int r = 0; r <= lastUsedRow; r++) {
            List<String> row = raw.get(r);
            List<String> padded = new ArrayList<>(width);
            for (int c = 0; c < width; c++) {
                String v = row != null && c < row.size() ? row.get(c) : "";
                padded.add(v != null ? v : "");
            }
            out.add(List.copyOf(padded));
        }
        return List.copyOf(out);
    }

    public record SheetGrid(String sheetName, List<List<String>> rows) {
        public SheetGrid {
            sheetName = sheetName != null ? sheetName : "";
            List<List<String>> copy = new ArrayList<>();
            if (rows != null) {
                for (List<String> row : rows) {
                    List<String> cells = new ArrayList<>();
                    if (row != null) {
                        for (String c : row) {
                            cells.add(c != null ? c : "");
                        }
                    }
                    copy.add(List.copyOf(cells));
                }
            }
            rows = List.copyOf(copy);
        }

        public static SheetGrid empty(String sheetName) {
            return new SheetGrid(sheetName, List.of());
        }
    }
}
