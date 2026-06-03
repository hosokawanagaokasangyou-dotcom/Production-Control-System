package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.io.PostProcessingProductMasterIo;

/**
 * {@code 後加工商品マスタ.xlsx} の1回読込キャッシュ。検索のたびに全件 POI 読込しない。
 */
final class PostProcessingProductMasterReferenceCache {

    private static volatile Snapshot cached;

    private PostProcessingProductMasterReferenceCache() {}

    record Snapshot(
            Path path,
            long lastModified,
            List<String> headers,
            List<ProductInfo> catalog,
            Map<String, Map<String, String>> rowByShohinCode) {}

    static Snapshot snapshot(Path referencePath) throws IOException {
        if (!Files.isRegularFile(referencePath)) {
            return emptySnapshot(referencePath);
        }
        long mtime = Files.getLastModifiedTime(referencePath).toMillis();
        Snapshot hit = cached;
        if (hit != null
                && Objects.equals(hit.path(), referencePath.toAbsolutePath().normalize())
                && hit.lastModified() == mtime) {
            return hit;
        }
        synchronized (PostProcessingProductMasterReferenceCache.class) {
            hit = cached;
            if (hit != null
                    && Objects.equals(hit.path(), referencePath.toAbsolutePath().normalize())
                    && hit.lastModified() == mtime) {
                return hit;
            }
            Snapshot loaded = load(referencePath, mtime);
            cached = loaded;
            return loaded;
        }
    }

    static void invalidate() {
        cached = null;
    }

    static Map<String, String> rowByCode(Path referencePath, String shohinCode) throws IOException {
        String key = PostProcessingProductMasterSearch.normalize(shohinCode);
        if (key.isEmpty()) {
            return Map.of();
        }
        return snapshot(referencePath).rowByShohinCode().getOrDefault(key, Map.of());
    }

    private static Snapshot load(Path referencePath, long mtime) throws IOException {
        Path abs = referencePath.toAbsolutePath().normalize();
        PlanInputTabularIo.TabularSheet sheet =
                PlanInputTabularIo.read(abs, PostProcessingProductMasterIo.DEFAULT_SHEET_NAME);
        List<String> headers = List.copyOf(sheet.headers());
        List<ProductInfo> catalog = new ArrayList<>();
        Map<String, Map<String, String>> rowByCode = new LinkedHashMap<>();
        for (List<String> row : sheet.rows()) {
            Map<String, String> map = PostProcessingProductMasterIo.rowToMap(headers, row);
            String code = map.getOrDefault("商品コード", "").trim();
            if (code.isEmpty()) {
                continue;
            }
            String norm = PostProcessingProductMasterSearch.normalize(code);
            catalog.add(toProductInfo(map));
            rowByCode.put(norm, map);
        }
        return new Snapshot(abs, mtime, headers, List.copyOf(catalog), Map.copyOf(rowByCode));
    }

    private static Snapshot emptySnapshot(Path referencePath) {
        Path abs =
                referencePath != null
                        ? referencePath.toAbsolutePath().normalize()
                        : Path.of("");
        return new Snapshot(abs, -1L, List.of(), List.of(), Map.of());
    }

    private static ProductInfo toProductInfo(Map<String, String> map) {
        return new ProductInfo(
                map.getOrDefault("商品コード", ""),
                map.getOrDefault("製品コード", ""),
                map.getOrDefault("商品名1", ""),
                map.getOrDefault("商品名2", ""),
                map.getOrDefault("単位名", ""),
                map.getOrDefault("入数", ""),
                map.getOrDefault("自社後加工区分", ""),
                map.getOrDefault("発泡体品名", ""),
                map.getOrDefault("発泡体品番", ""),
                map.getOrDefault("発泡体幅", ""),
                map.getOrDefault("発泡体長さ", ""),
                map.getOrDefault("発泡体色", ""),
                map.getOrDefault("発泡体厚み", ""),
                "");
    }
}
