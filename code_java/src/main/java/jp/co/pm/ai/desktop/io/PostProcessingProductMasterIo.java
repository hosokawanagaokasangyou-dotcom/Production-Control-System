package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.reconciliation.PostProcessingProductMasterColumnGroups;
import jp.co.pm.ai.desktop.reconciliation.PostProcessingProductMasterSearch;

/**
 * {@code 後加工商品マスタ.xlsx} およびアップロード用コピーの読み書き。
 */
public final class PostProcessingProductMasterIo {

    public static final String REFERENCE_FILE_NAME = "後加工商品マスタ.xlsx";
    public static final String DEFAULT_UPLOAD_FILE_NAME = "アップロード用_後加工商品マスタ.xlsx";
    public static final String DEFAULT_SHEET_NAME = "Sheet1";
    public static final int DEFAULT_SEARCH_LIMIT = 200;

    private PostProcessingProductMasterIo() {}

    public record SearchHit(
            String shohinCode,
            String shohinName1,
            String foamPartNo,
            String foamName,
            Map<String, String> rowByColumn) {}

    public record SearchFilter(
            String shohinCode,
            String foamPartNo,
            String foamType,
            String foamLength,
            String foamName) {

        public static SearchFilter empty() {
            return new SearchFilter("", "", "", "", "");
        }
    }

    public static Path resolveReferencePath(Map<String, String> ui) {
        return AppPaths.resolveAladdinMasterDir(ui).resolve(REFERENCE_FILE_NAME);
    }

    public static Path resolveUploadPath(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String override = trimEnv(u.get(AppPaths.KEY_PM_AI_POSTPROC_PRODUCT_MASTER_UPLOAD));
        if (!override.isEmpty()) {
            return Path.of(override).toAbsolutePath().normalize();
        }
        return AppPaths.resolveAladdinMasterDir(ui).resolve(DEFAULT_UPLOAD_FILE_NAME);
    }

    public static List<String> readHeaders(Path path) throws IOException {
        PlanInputTabularIo.TabularRead read =
                PlanInputTabularIo.readWithResolvedSheet(path, DEFAULT_SHEET_NAME);
        return List.copyOf(read.tabular().headers());
    }

    public static List<SearchHit> searchReference(
            Path referencePath,
            SearchFilter filter,
            int limit,
            java.util.List<jp.co.pm.ai.desktop.reconciliation.ProductInfo> integratedCatalog)
            throws IOException {
        return PostProcessingProductMasterSearch.searchReference(
                referencePath, filter, limit, integratedCatalog);
    }

    public static List<SearchHit> searchReference(
            Path referencePath, SearchFilter filter, int limit) throws IOException {
        return searchReference(referencePath, filter, limit, java.util.List.of());
    }

    public static Map<String, String> loadRowByShohinCode(Path path, String shohinCode)
            throws IOException {
        return PostProcessingProductMasterSearch.loadRowByShohinCode(path, shohinCode);
    }

    public static PlanInputTabularIo.TabularSheet readUploadWorkbook(Path path) throws IOException {
        if (!Files.isRegularFile(path)) {
            return new PlanInputTabularIo.TabularSheet(List.of(), List.of());
        }
        return PlanInputTabularIo.read(path, DEFAULT_SHEET_NAME);
    }

    public static void writeUploadWorkbook(
            Path path, List<String> headers, List<Map<String, String>> rows) throws IOException {
        List<String> hdr = PostProcessingProductMasterColumnGroups.alignHeadersToReference(headers);
        List<List<String>> dataRows = new ArrayList<>();
        for (Map<String, String> row : rows) {
            List<String> line = new ArrayList<>(hdr.size());
            for (String h : hdr) {
                line.add(row != null ? row.getOrDefault(h, "") : "");
            }
            dataRows.add(line);
        }
        Path parent = path.getParent();
        if (parent != null && !Files.isDirectory(parent)) {
            Files.createDirectories(parent);
        }
        PlanInputTabularIo.write(
                path, DEFAULT_SHEET_NAME, new PlanInputTabularIo.TabularSheet(hdr, dataRows));
    }

    public static void createEmptyUploadFromReference(Path referencePath, Path uploadPath)
            throws IOException {
        List<String> headers = readHeaders(referencePath);
        writeUploadWorkbook(uploadPath, headers, List.of());
    }

    public static Map<String, String> rowToMap(List<String> headers, List<String> row) {
        Map<String, String> map = new LinkedHashMap<>();
        for (int i = 0; i < headers.size(); i++) {
            String key = headers.get(i) != null ? headers.get(i).trim() : "";
            if (key.isEmpty()) {
                continue;
            }
            String val = i < row.size() && row.get(i) != null ? row.get(i).trim() : "";
            map.put(key, val);
        }
        return map;
    }

    private static String trimEnv(String s) {
        return s != null ? s.strip() : "";
    }
}
