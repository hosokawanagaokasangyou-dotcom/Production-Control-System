package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.CodeDispatchLookupTableIo;
import jp.co.pm.ai.desktop.io.CodeDispatchLookupTableIo.KeyValTable;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.planning.stage2.core.Stage2RollUnitLengthTables;

/**
 * 段階1出力 {@code plan_input_tasks.xlsx} に現れた製品名・使用原反を、{@code code/} 配下の参照テーブルへ不足分のみ追記する。
 */
public final class CodeDispatchLookupTablesMerge {

    public static final String FILE_USED_RAW_ROLL = "使用原反,ロール単位の長さ.txt";
    public static final String FILE_PRODUCT_ROLL = "製品名,ロール単位の長さ.txt";
    public static final String FILE_PRODUCT_WIDTH = "製品名, 製品幅.txt";
    public static final String FILE_PRODUCT_THICK = "製品名,製品厚み.txt";
    public static final String FILE_PRODUCT_LENGTH = "製品名,製品長.txt";
    public static final String FILE_USED_RAW_WIDTH = "使用原反, 加工幅.txt";

    private static final String HDR_USED_RAW_ROLL = "使用原反,ロール単位の長さ";
    private static final String HDR_PRODUCT_ROLL = "製品名,ロール単位の長さ";
    private static final String HDR_PRODUCT_WIDTH = "製品名,製品幅";
    private static final String HDR_PRODUCT_THICK = "製品名,製品厚み";
    private static final String HDR_PRODUCT_LENGTH = "製品名,製品長";
    private static final String HDR_USED_RAW_WIDTH = "使用原反,原反幅";

    private static final String COL_PRODUCT = "製品名";
    private static final String COL_USED_RAW = "使用原反";
    private static final String COL_ROLL_M = "ロール単位長さ";
    private static final String COL_WIDTH = "製品幅";
    private static final String COL_THICK = "製品厚み";
    private static final String COL_LENGTH = "製品長";
    private static final String COL_RAW_ROLL = "(原反)ロール単位長さ";
    private static final String COL_RAW_ROLL_ALT = "（原反）ロール単位長さ";
    private static final String COL_RAW_WIDTH = "原反幅";

    public record MergeSummary(
            int addedProductRoll,
            int addedProductWidth,
            int addedProductThick,
            int addedProductLength,
            int addedUsedRawRoll,
            int addedUsedRawWidth) {

        public int totalAdded() {
            return addedProductRoll
                    + addedProductWidth
                    + addedProductThick
                    + addedProductLength
                    + addedUsedRawRoll
                    + addedUsedRawWidth;
        }

        public String summaryJa() {
            if (totalAdded() <= 0) {
                return "追記なし";
            }
            return "製品ロール "
                    + addedProductRoll
                    + "、製品幅 "
                    + addedProductWidth
                    + "、厚み "
                    + addedProductThick
                    + "、製品長 "
                    + addedProductLength
                    + "、原反ロール "
                    + addedUsedRawRoll
                    + "、原反幅 "
                    + addedUsedRawWidth;
        }
    }

    private CodeDispatchLookupTablesMerge() {}

    public static MergeSummary mergeAfterStage1(Map<String, String> ui) throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path plan = AppPaths.defaultStage1PlanTasksPath(u);
        if (!Files.isRegularFile(plan)) {
            return new MergeSummary(0, 0, 0, 0, 0, 0);
        }
        PlanInputTabularIo.TabularRead tr =
                PlanInputTabularIo.readWithResolvedSheet(plan, AppPaths.STAGE1_PLAN_OUTPUT_SHEET);
        List<String> headers = tr.tabular().headers();
        List<List<String>> rows = tr.tabular().rows();
        if (headers == null || headers.isEmpty()) {
            return new MergeSummary(0, 0, 0, 0, 0, 0);
        }
        Path codeDir = AppPaths.resolveRepoRoot(u).resolve("code");

        int apRoll = 0;
        int apW = 0;
        int apT = 0;
        int apL = 0;
        int auRoll = 0;
        int auW = 0;

        int iProd = headers.indexOf(COL_PRODUCT);
        int iUsed = headers.indexOf(COL_USED_RAW);
        int iRoll = headers.indexOf(COL_ROLL_M);
        int iPw = headers.indexOf(COL_WIDTH);
        int iPt = headers.indexOf(COL_THICK);
        int iPl = headers.indexOf(COL_LENGTH);
        int iRawRoll = headers.indexOf(COL_RAW_ROLL);
        if (iRawRoll < 0) {
            iRawRoll = headers.indexOf(COL_RAW_ROLL_ALT);
        }
        int iRawW = headers.indexOf(COL_RAW_WIDTH);

        if (iProd >= 0) {
            Path p = codeDir.resolve(FILE_PRODUCT_ROLL);
            KeyValTable cur = CodeDispatchLookupTableIo.readOrEmpty(p, HDR_PRODUCT_ROLL);
            LinkedHashMap<String, String> m = new LinkedHashMap<>(cur.rows());
            for (List<String> row : rows) {
                String prod = cell(row, iProd);
                if (prod.isBlank()) {
                    continue;
                }
                if (containsKeyNormalized(m, prod)) {
                    continue;
                }
                double rollCell = iRoll >= 0 ? Stage2RollUnitLengthTables.parseFloatSafe(cell(row, iRoll), 0.0) : 0.0;
                double v;
                if (rollCell > 1e-12) {
                    v = rollCell;
                } else {
                    v = Stage2RollUnitLengthTables.inferFromProductDimensions(prod, 100.0);
                }
                m.put(prod.strip(), formatNum(v));
                apRoll++;
            }
            if (apRoll > 0) {
                CodeDispatchLookupTableIo.write(p, new KeyValTable(cur.headerLine(), m));
            }
        }
        if (iProd >= 0 && iPw >= 0) {
            Path p = codeDir.resolve(FILE_PRODUCT_WIDTH);
            KeyValTable cur = CodeDispatchLookupTableIo.readOrEmpty(p, HDR_PRODUCT_WIDTH);
            LinkedHashMap<String, String> m = new LinkedHashMap<>(cur.rows());
            for (List<String> row : rows) {
                String prod = cell(row, iProd);
                if (prod.isBlank()) {
                    continue;
                }
                if (containsKeyNormalized(m, prod)) {
                    continue;
                }
                double w = Stage2RollUnitLengthTables.parseFloatSafe(cell(row, iPw), 0.0);
                if (!(w > 1e-12)) {
                    continue;
                }
                m.put(prod.strip(), formatNum(w));
                apW++;
            }
            if (apW > 0) {
                CodeDispatchLookupTableIo.write(p, new KeyValTable(cur.headerLine(), m));
            }
        }
        if (iProd >= 0 && iPt >= 0) {
            Path p = codeDir.resolve(FILE_PRODUCT_THICK);
            KeyValTable cur = CodeDispatchLookupTableIo.readOrEmpty(p, HDR_PRODUCT_THICK);
            LinkedHashMap<String, String> m = new LinkedHashMap<>(cur.rows());
            for (List<String> row : rows) {
                String prod = cell(row, iProd);
                if (prod.isBlank()) {
                    continue;
                }
                if (containsKeyNormalized(m, prod)) {
                    continue;
                }
                double t = Stage2RollUnitLengthTables.parseFloatSafe(cell(row, iPt), 0.0);
                if (!(t > 1e-12)) {
                    continue;
                }
                m.put(prod.strip(), formatNum(t));
                apT++;
            }
            if (apT > 0) {
                CodeDispatchLookupTableIo.write(p, new KeyValTable(cur.headerLine(), m));
            }
        }
        if (iProd >= 0 && iPl >= 0) {
            Path p = codeDir.resolve(FILE_PRODUCT_LENGTH);
            KeyValTable cur = CodeDispatchLookupTableIo.readOrEmpty(p, HDR_PRODUCT_LENGTH);
            LinkedHashMap<String, String> m = new LinkedHashMap<>(cur.rows());
            for (List<String> row : rows) {
                String prod = cell(row, iProd);
                if (prod.isBlank()) {
                    continue;
                }
                if (containsKeyNormalized(m, prod)) {
                    continue;
                }
                double len = Stage2RollUnitLengthTables.parseFloatSafe(cell(row, iPl), 0.0);
                if (!(len > 1e-12)) {
                    continue;
                }
                m.put(prod.strip(), formatNum(len));
                apL++;
            }
            if (apL > 0) {
                CodeDispatchLookupTableIo.write(p, new KeyValTable(cur.headerLine(), m));
            }
        }
        if (iUsed >= 0) {
            Path p = codeDir.resolve(FILE_USED_RAW_ROLL);
            KeyValTable cur = CodeDispatchLookupTableIo.readOrEmpty(p, HDR_USED_RAW_ROLL);
            LinkedHashMap<String, String> m = new LinkedHashMap<>(cur.rows());
            for (List<String> row : rows) {
                String ur = cell(row, iUsed);
                if (ur.isBlank()) {
                    continue;
                }
                if (containsKeyNormalized(m, ur)) {
                    continue;
                }
                double rawRoll =
                        iRawRoll >= 0 ? Stage2RollUnitLengthTables.parseFloatSafe(cell(row, iRawRoll), 0.0) : 0.0;
                if (!(rawRoll > 1e-12)) {
                    continue;
                }
                m.put(ur.strip(), formatNum(rawRoll));
                auRoll++;
            }
            if (auRoll > 0) {
                CodeDispatchLookupTableIo.write(p, new KeyValTable(cur.headerLine(), m));
            }
        }
        if (iUsed >= 0 && iRawW >= 0) {
            Path p = codeDir.resolve(FILE_USED_RAW_WIDTH);
            KeyValTable cur = CodeDispatchLookupTableIo.readOrEmpty(p, HDR_USED_RAW_WIDTH);
            LinkedHashMap<String, String> m = new LinkedHashMap<>(cur.rows());
            for (List<String> row : rows) {
                String ur = cell(row, iUsed);
                if (ur.isBlank()) {
                    continue;
                }
                if (containsKeyNormalized(m, ur)) {
                    continue;
                }
                double rw = Stage2RollUnitLengthTables.parseFloatSafe(cell(row, iRawW), 0.0);
                if (!(rw > 1e-12)) {
                    continue;
                }
                m.put(ur.strip(), formatNum(rw));
                auW++;
            }
            if (auW > 0) {
                CodeDispatchLookupTableIo.write(p, new KeyValTable(cur.headerLine(), m));
            }
        }
        return new MergeSummary(apRoll, apW, apT, apL, auRoll, auW);
    }

    private static String cell(List<String> row, int col) {
        if (row == null || col < 0 || col >= row.size()) {
            return "";
        }
        String v = row.get(col);
        return v != null ? v.strip() : "";
    }

    private static boolean containsKeyNormalized(LinkedHashMap<String, String> m, String key) {
        String nk = Stage2RollUnitLengthTables.normalizeKey(key);
        if (nk.isEmpty()) {
            return true;
        }
        for (String k : m.keySet()) {
            if (nk.equals(Stage2RollUnitLengthTables.normalizeKey(k))) {
                return true;
            }
        }
        return false;
    }

    private static String formatNum(double v) {
        if (!Double.isFinite(v)) {
            return "";
        }
        long r = Math.round(v);
        if (Math.abs(v - r) < 1e-6) {
            return Long.toString(r);
        }
        return String.format(Locale.ROOT, "%s", v);
    }
}
