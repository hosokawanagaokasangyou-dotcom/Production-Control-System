package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.io.PostProcessingProductMasterIo;

/** 参照マスタ xlsx のフィルタ検索（{@link RequestFormMasterProductCandidateMatcher} 再利用）。 */
public final class PostProcessingProductMasterSearch {

    private PostProcessingProductMasterSearch() {}

    public static List<PostProcessingProductMasterIo.SearchHit> searchReference(
            Path referencePath, PostProcessingProductMasterIo.SearchFilter filter, int limit)
            throws IOException {
        if (!Files.isRegularFile(referencePath)) {
            return List.of();
        }
        int cap = limit > 0 ? limit : PostProcessingProductMasterIo.DEFAULT_SEARCH_LIMIT;
        PlanInputTabularIo.TabularSheet sheet =
                PlanInputTabularIo.read(referencePath, PostProcessingProductMasterIo.DEFAULT_SHEET_NAME);
        List<String> headers = sheet.headers();
        if (headers.isEmpty()) {
            return List.of();
        }
        PostProcessingProductMasterIo.SearchFilter f =
                filter != null ? filter : PostProcessingProductMasterIo.SearchFilter.empty();

        String kwCode = normalize(filterKeyword(f.shohinCode()));
        String kwPart = normalize(filterKeyword(f.foamPartNo()));
        String kwType = normalize(filterKeyword(f.foamType()));
        String kwLength = normalizeLength(filterKeyword(f.foamLength()));
        String kwName = normalize(filterKeyword(f.foamName()));

        boolean anyKeyword =
                !kwCode.isEmpty()
                        || !kwPart.isEmpty()
                        || !kwType.isEmpty()
                        || !kwLength.isEmpty()
                        || !kwName.isEmpty();

        List<ProductInfo> catalog = new ArrayList<>();
        List<Map<String, String>> rowMaps = new ArrayList<>();
        for (List<String> row : sheet.rows()) {
            Map<String, String> map = PostProcessingProductMasterIo.rowToMap(headers, row);
            String code = map.getOrDefault("商品コード", "").trim();
            if (code.isEmpty()) {
                continue;
            }
            rowMaps.add(map);
            catalog.add(toProductInfo(map));
        }

        List<PostProcessingProductMasterIo.SearchHit> hits = new ArrayList<>();
        if (!anyKeyword) {
            for (int i = 0; i < rowMaps.size() && hits.size() < cap; i++) {
                hits.add(toHit(rowMaps.get(i)));
            }
            return hits;
        }

        List<String> labels =
                RequestFormMasterProductCandidateMatcher.buildRankedCandidateLabels(
                        catalog, kwCode, kwPart, kwType, kwLength, kwName, cap);
        for (String label : labels) {
            String code = labelCodeFromCandidateLabel(label);
            for (Map<String, String> map : rowMaps) {
                if (code.equals(normalize(map.getOrDefault("商品コード", "")))) {
                    hits.add(toHit(map));
                    break;
                }
            }
        }
        return hits;
    }

    public static String normalize(String val) {
        return RequestFormMasterProductCandidateMatcher.normalize(val);
    }

    private static String normalizeLength(String val) {
        return RequestFormMasterProductCandidateMatcher.normalizeLengthKeyword(val);
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

    private static PostProcessingProductMasterIo.SearchHit toHit(Map<String, String> map) {
        return new PostProcessingProductMasterIo.SearchHit(
                map.getOrDefault("商品コード", ""),
                map.getOrDefault("商品名1", ""),
                map.getOrDefault("発泡体品番", ""),
                map.getOrDefault("発泡体品名", ""),
                Map.copyOf(map));
    }

    private static String labelCodeFromCandidateLabel(String label) {
        if (label == null || label.isBlank()) {
            return "";
        }
        int sep = label.indexOf(" | ");
        String code = sep >= 0 ? label.substring(0, sep) : label;
        return normalize(code);
    }

    private static String filterKeyword(String s) {
        return s != null ? s.trim() : "";
    }
}
