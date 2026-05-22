package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
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
 * 段階1後に {@code code/} 材料テーブルへ空欄追記されたキーを集約し、ダイアログ入力結果を書き戻す。
 */
public final class CodeDispatchLookupTablesBlankPrompt {

    private static final String COL_PRODUCT = "製品名";
    private static final String COL_USED_RAW = "使用原反";
    private static final String COL_ROLL_M = "(製品)ロール単位長さ";
    private static final String COL_ROLL_M_LEGACY = "ロール単位長さ";
    private static final String COL_WIDTH = "製品幅";
    private static final String COL_THICK = "製品厚み";
    private static final String COL_LENGTH = "製品長";
    private static final String COL_RAW_ROLL = "(原反)ロール単位長さ";
    private static final String COL_RAW_ROLL_ALT = "（原反）ロール単位長さ";
    private static final String COL_RAW_WIDTH = "原反幅";

    private static final String LBL_PRODUCT_ROLL = "製品名→ロール長(m)";
    private static final String LBL_PRODUCT_WIDTH = "製品名→製品幅(mm)";
    private static final String LBL_PRODUCT_THICK = "製品名→厚み(mm)";
    private static final String LBL_PRODUCT_LENGTH = "製品名→製品長(mm)";
    private static final String LBL_USED_RAW_ROLL = "使用原反→ロール長(m)";
    private static final String LBL_USED_RAW_WIDTH = "使用原反→原反幅(mm)";

    public record ProductPromptRow(
            String productName,
            boolean needRollLength,
            boolean needWidth,
            boolean needThickness,
            boolean needLength,
            String suggestedRollLength,
            String suggestedWidth,
            String suggestedThickness,
            String suggestedLength) {}

    public record UsedRawPromptRow(
            String usedRaw,
            boolean needRollLength,
            boolean needRawWidth,
            String suggestedRollLength,
            String suggestedRawWidth) {}

    public record PromptBundle(List<ProductPromptRow> products, List<UsedRawPromptRow> usedRaws) {
        public boolean empty() {
            return (products == null || products.isEmpty())
                    && (usedRaws == null || usedRaws.isEmpty());
        }
    }

    public record ProductInput(
            String productName,
            String rollLength,
            String productWidth,
            String thickness,
            String productLength) {}

    public record UsedRawInput(String usedRaw, String rollLength, String rawWidth) {}

  public record ApplySummary(int updatedFields) {}

    private CodeDispatchLookupTablesBlankPrompt() {}

    public static PromptBundle collectPrompt(
            Map<String, String> ui, CodeDispatchLookupTablesValidator.ValidationResult vr)
            throws IOException {
        if (vr == null || vr.ok()) {
            return new PromptBundle(List.of(), List.of());
        }
        Map<String, String> u = ui != null ? ui : Map.of();
        PlanHints hints = loadPlanHints(u);

        LinkedHashMap<String, ProductFlags> products = new LinkedHashMap<>();
        LinkedHashMap<String, UsedRawFlags> usedRaws = new LinkedHashMap<>();

        for (CodeDispatchLookupTablesValidator.BlankValueIssue issue : vr.issues()) {
            if (issue == null) {
                continue;
            }
            String label = issue.tableLabelJa() != null ? issue.tableLabelJa() : "";
            String key = issue.key() != null ? issue.key().strip() : "";
            if (key.isEmpty()) {
                continue;
            }
            switch (label) {
                case LBL_PRODUCT_ROLL ->
                        products.computeIfAbsent(key, ProductFlags::new).needRoll = true;
                case LBL_PRODUCT_WIDTH ->
                        products.computeIfAbsent(key, ProductFlags::new).needWidth = true;
                case LBL_PRODUCT_THICK ->
                        products.computeIfAbsent(key, ProductFlags::new).needThickness = true;
                case LBL_PRODUCT_LENGTH ->
                        products.computeIfAbsent(key, ProductFlags::new).needLength = true;
                case LBL_USED_RAW_ROLL ->
                        usedRaws.computeIfAbsent(key, UsedRawFlags::new).needRoll = true;
                case LBL_USED_RAW_WIDTH ->
                        usedRaws.computeIfAbsent(key, UsedRawFlags::new).needWidth = true;
                default -> {
                    // ignore unknown labels
                }
            }
        }

        List<ProductPromptRow> productRows = new ArrayList<>(products.size());
        for (ProductFlags pf : products.values()) {
            PlanProductHint ph = hints.products().get(normalize(pf.key));
            productRows.add(
                    new ProductPromptRow(
                            pf.key,
                            pf.needRoll,
                            pf.needWidth,
                            pf.needThickness,
                            pf.needLength,
                            suggest(
                                    pf.needRoll,
                                    ph != null ? ph.rollLength() : "",
                                    inferRollFromName(pf.key)),
                            suggest(
                                    pf.needWidth,
                                    ph != null ? ph.width() : "",
                                    inferWidthFromName(pf.key)),
                            suggest(
                                    pf.needThickness,
                                    ph != null ? ph.thickness() : "",
                                    ""),
                            suggest(
                                    pf.needLength,
                                    ph != null ? ph.length() : "",
                                    inferRollFromName(pf.key))));
        }

        List<UsedRawPromptRow> usedRawRows = new ArrayList<>(usedRaws.size());
        for (UsedRawFlags uf : usedRaws.values()) {
            PlanUsedRawHint uh = hints.usedRaws().get(normalize(uf.key));
            usedRawRows.add(
                    new UsedRawPromptRow(
                            uf.key,
                            uf.needRoll,
                            uf.needWidth,
                            suggest(
                                    uf.needRoll,
                                    uh != null ? uh.rollLength() : "",
                                    inferRollFromName(uf.key)),
                            suggest(
                                    uf.needWidth,
                                    uh != null ? uh.rawWidth() : "",
                                    inferWidthFromName(uf.key))));
        }
        return new PromptBundle(List.copyOf(productRows), List.copyOf(usedRawRows));
    }

    public static ApplySummary applyInputs(
            Map<String, String> ui, List<ProductInput> products, List<UsedRawInput> usedRaws)
            throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        AppPaths.ensureAllDispatchLookupTablesFromRepoIfMissing(u);
        int updated = 0;

        if (products != null) {
            for (ProductInput in : products) {
                if (in == null || in.productName() == null || in.productName().isBlank()) {
                    continue;
                }
                String key = in.productName().strip();
                if (in.rollLength() != null && !in.rollLength().isBlank()) {
                    updated +=
                            updateTableValue(
                                    AppPaths.dispatchLookupTablePath(u, CodeDispatchLookupTablesMerge.FILE_PRODUCT_ROLL),
                                    CodeDispatchLookupTablesMerge.FILE_PRODUCT_ROLL.replace(".txt", ""),
                                    key,
                                    in.rollLength().strip());
                }
                if (in.productWidth() != null && !in.productWidth().isBlank()) {
                    updated +=
                            updateTableValue(
                                    AppPaths.dispatchLookupTablePath(u, CodeDispatchLookupTablesMerge.FILE_PRODUCT_WIDTH),
                                    "製品名,製品幅",
                                    key,
                                    in.productWidth().strip());
                }
                if (in.thickness() != null && !in.thickness().isBlank()) {
                    updated +=
                            updateTableValue(
                                    AppPaths.dispatchLookupTablePath(u, CodeDispatchLookupTablesMerge.FILE_PRODUCT_THICK),
                                    "製品名,製品厚み",
                                    key,
                                    in.thickness().strip());
                }
                if (in.productLength() != null && !in.productLength().isBlank()) {
                    updated +=
                            updateTableValue(
                                    AppPaths.dispatchLookupTablePath(u, CodeDispatchLookupTablesMerge.FILE_PRODUCT_LENGTH),
                                    "製品名,製品長",
                                    key,
                                    in.productLength().strip());
                }
            }
        }
        if (usedRaws != null) {
            for (UsedRawInput in : usedRaws) {
                if (in == null || in.usedRaw() == null || in.usedRaw().isBlank()) {
                    continue;
                }
                String key = in.usedRaw().strip();
                if (in.rollLength() != null && !in.rollLength().isBlank()) {
                    updated +=
                            updateTableValue(
                                    AppPaths.dispatchLookupTablePath(u, CodeDispatchLookupTablesMerge.FILE_USED_RAW_ROLL),
                                    CodeDispatchLookupTablesMerge.FILE_USED_RAW_ROLL.replace(".txt", ""),
                                    key,
                                    in.rollLength().strip());
                }
                if (in.rawWidth() != null && !in.rawWidth().isBlank()) {
                    updated +=
                            updateTableValue(
                                    AppPaths.dispatchLookupTablePath(u, CodeDispatchLookupTablesMerge.FILE_USED_RAW_WIDTH),
                                    "使用原反,原反幅",
                                    key,
                                    in.rawWidth().strip());
                }
            }
        }
        return new ApplySummary(updated);
    }

    private static int updateTableValue(Path path, String defaultHeader, String key, String value)
            throws IOException {
        KeyValTable cur = CodeDispatchLookupTableIo.readOrEmpty(path, defaultHeader);
        LinkedHashMap<String, String> rows = new LinkedHashMap<>(cur.rows());
        String nk = normalize(key);
        if (nk.isEmpty()) {
            return 0;
        }
        String actualKey = null;
        for (String k : rows.keySet()) {
            if (nk.equals(normalize(k))) {
                actualKey = k;
                break;
            }
        }
        if (actualKey == null) {
            actualKey = key.strip();
            rows.put(actualKey, value);
            CodeDispatchLookupTableIo.write(path, new KeyValTable(cur.headerLine(), rows));
            return 1;
        }
        String curVal = rows.get(actualKey);
        if (curVal != null && !curVal.strip().isEmpty()) {
            return 0;
        }
        rows.put(actualKey, value);
        CodeDispatchLookupTableIo.write(path, new KeyValTable(cur.headerLine(), rows));
        return 1;
    }

    private static String suggest(boolean needed, String fromPlan, String inferred) {
        if (!needed) {
            return "";
        }
        if (fromPlan != null && !fromPlan.isBlank()) {
            return fromPlan.strip();
        }
        if (inferred != null && !inferred.isBlank()) {
            return inferred.strip();
        }
        return "";
    }

    private static String inferRollFromName(String name) {
        double v = Stage2RollUnitLengthTables.inferFromProductDimensions(name, 0.0);
        return v > 1e-12 ? formatNum(v) : "";
    }

    private static String inferWidthFromName(String name) {
        if (name == null || name.isBlank()) {
            return "";
        }
        String s = Stage2RollUnitLengthTables.normalizeKey(name);
        java.util.regex.Pattern pair =
                java.util.regex.Pattern.compile("(\\d{2,6})[xX](\\d{2,6})");
        java.util.regex.Matcher mp = pair.matcher(s);
        int lastA = -1;
        while (mp.find()) {
            try {
                lastA = Integer.parseInt(mp.group(1));
            } catch (NumberFormatException ignored) {
                lastA = -1;
            }
        }
        return lastA > 0 ? Integer.toString(lastA) : "";
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

    private static String normalize(String key) {
        return Stage2RollUnitLengthTables.normalizeKey(key);
    }

    private record PlanHints(
            Map<String, PlanProductHint> products, Map<String, PlanUsedRawHint> usedRaws) {}

    private record PlanProductHint(
            String rollLength, String width, String thickness, String length) {}

    private record PlanUsedRawHint(String rollLength, String rawWidth) {}

    private static PlanHints loadPlanHints(Map<String, String> ui) throws IOException {
        Path plan = AppPaths.defaultStage1PlanTasksPath(ui);
        if (!java.nio.file.Files.isRegularFile(plan)) {
            return new PlanHints(Map.of(), Map.of());
        }
        PlanInputTabularIo.TabularRead tr =
                PlanInputTabularIo.readWithResolvedSheet(plan, AppPaths.STAGE1_PLAN_OUTPUT_SHEET);
        List<String> headers = tr.tabular().headers();
        List<List<String>> rows = tr.tabular().rows();
        if (headers == null || headers.isEmpty() || rows == null) {
            return new PlanHints(Map.of(), Map.of());
        }
        int iProd = headers.indexOf(COL_PRODUCT);
        int iUsed = headers.indexOf(COL_USED_RAW);
        int iRoll = headers.indexOf(COL_ROLL_M);
        if (iRoll < 0) {
            iRoll = headers.indexOf(COL_ROLL_M_LEGACY);
        }
        int iPw = headers.indexOf(COL_WIDTH);
        int iPt = headers.indexOf(COL_THICK);
        int iPl = headers.indexOf(COL_LENGTH);
        int iRawRoll = headers.indexOf(COL_RAW_ROLL);
        if (iRawRoll < 0) {
            iRawRoll = headers.indexOf(COL_RAW_ROLL_ALT);
        }
        int iRawW = headers.indexOf(COL_RAW_WIDTH);

        LinkedHashMap<String, PlanProductHint> products = new LinkedHashMap<>();
        LinkedHashMap<String, PlanUsedRawHint> usedRaws = new LinkedHashMap<>();
        for (List<String> row : rows) {
            if (iProd >= 0) {
                String prod = cell(row, iProd);
                if (!prod.isBlank()) {
                    String nk = normalize(prod);
                    products.putIfAbsent(
                            nk,
                            new PlanProductHint(
                                    cellPositive(row, iRoll),
                                    cellPositive(row, iPw),
                                    cellPositive(row, iPt),
                                    cellPositive(row, iPl)));
                }
            }
            if (iUsed >= 0) {
                String ur = cell(row, iUsed);
                if (!ur.isBlank()) {
                    String nk = normalize(ur);
                    usedRaws.putIfAbsent(
                            nk,
                            new PlanUsedRawHint(cellPositive(row, iRawRoll), cellPositive(row, iRawW)));
                }
            }
        }
        return new PlanHints(Map.copyOf(products), Map.copyOf(usedRaws));
    }

    private static String cell(List<String> row, int col) {
        if (row == null || col < 0 || col >= row.size()) {
            return "";
        }
        String v = row.get(col);
        return v != null ? v.strip() : "";
    }

    private static String cellPositive(List<String> row, int col) {
        String v = cell(row, col);
        if (v.isBlank() || "不明".equals(v)) {
            return "";
        }
        double n = Stage2RollUnitLengthTables.parseFloatSafe(v, 0.0);
        return n > 1e-12 ? formatNum(n) : "";
    }

    private static final class ProductFlags {
        final String key;
        boolean needRoll;
        boolean needWidth;
        boolean needThickness;
        boolean needLength;

        ProductFlags(String key) {
            this.key = key;
        }
    }

    private static final class UsedRawFlags {
        final String key;
        boolean needRoll;
        boolean needWidth;

        UsedRawFlags(String key) {
            this.key = key;
        }
    }
}
