package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import jp.co.pm.ai.desktop.io.PostProcessingProductMasterIo;

/** 参照マスタのフィルタ検索（{@link RequestFormMasterProductCandidateMatcher} と同一スコアリング）。 */
public final class PostProcessingProductMasterSearch {

    /** 後加工商品マスタ参照検索の先頭文字フィルタ（製品側・原反側）。 */
    public record MasterReferencePrefixFilters(List<String> productSide, List<String> rawSide) {

        public static MasterReferencePrefixFilters none() {
            return new MasterReferencePrefixFilters(List.of(), List.of());
        }
    }

    private PostProcessingProductMasterSearch() {}

    /**
     * 検索実行。
     *
     * @param referencePath 雛形の全154列を得る {@code 後加工商品マスタ.xlsx}
     * @param integratedCatalog 依頼書と同じメモリ上リスト（統合マスタ②）。空でなければ検索はこれを使い高速化する。
     */
    public static List<PostProcessingProductMasterIo.SearchHit> searchReference(
            Path referencePath,
            PostProcessingProductMasterIo.SearchFilter filter,
            int limit,
            List<ProductInfo> integratedCatalog)
            throws IOException {
        return searchReference(
                referencePath, filter, limit, integratedCatalog, MasterReferencePrefixFilters.none());
    }

    public static List<PostProcessingProductMasterIo.SearchHit> searchReference(
            Path referencePath,
            PostProcessingProductMasterIo.SearchFilter filter,
            int limit,
            List<ProductInfo> integratedCatalog,
            MasterReferencePrefixFilters prefixFilters)
            throws IOException {
        int cap = limit > 0 ? limit : PostProcessingProductMasterIo.DEFAULT_SEARCH_LIMIT;
        PostProcessingProductMasterIo.SearchFilter f =
                filter != null ? filter : PostProcessingProductMasterIo.SearchFilter.empty();
        MasterReferencePrefixFilters prefixes =
                prefixFilters != null ? prefixFilters : MasterReferencePrefixFilters.none();

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

        List<ProductInfo> rawCatalog;
        if (integratedCatalog != null && !integratedCatalog.isEmpty()) {
            rawCatalog = integratedCatalog;
        } else {
            rawCatalog = PostProcessingProductMasterReferenceCache.snapshot(referencePath).catalog();
        }
        List<ProductInfo> catalogForScore =
                RequestFormMasterProductCandidateMatcher.filterCatalogForMasterReferenceSearch(
                        rawCatalog, prefixes.productSide(), prefixes.rawSide());

        PostProcessingProductMasterReferenceCache.Snapshot refSnap =
                PostProcessingProductMasterReferenceCache.snapshot(referencePath);
        Map<String, Map<String, String>> rowsByCode = refSnap.rowByShohinCode();

        List<PostProcessingProductMasterIo.SearchHit> hits = new ArrayList<>();
        if (!anyKeyword) {
            for (int i = 0; i < catalogForScore.size() && hits.size() < cap; i++) {
                hits.add(toHit(catalogForScore.get(i), rowsByCode));
            }
            return hits;
        }

        List<String> labels =
                RequestFormMasterProductCandidateMatcher.buildRankedCandidateLabels(
                        catalogForScore, kwCode, kwPart, kwType, kwLength, kwName, cap);
        for (String label : labels) {
            String code = labelCodeFromCandidateLabel(label);
            Map<String, String> row = rowsByCode.get(code);
            if (row != null && !row.isEmpty()) {
                hits.add(
                        new PostProcessingProductMasterIo.SearchHit(
                                row.getOrDefault("商品コード", ""),
                                row.getOrDefault("商品名1", ""),
                                row.getOrDefault("発泡体品番", ""),
                                row.getOrDefault("発泡体品名", ""),
                                Map.copyOf(row)));
            } else {
                ProductInfo p = findByCode(catalogForScore, code);
                if (p != null) {
                    hits.add(toHit(p, rowsByCode));
                }
            }
        }
        return hits;
    }

    public static List<PostProcessingProductMasterIo.SearchHit> searchReference(
            Path referencePath, PostProcessingProductMasterIo.SearchFilter filter, int limit)
            throws IOException {
        return searchReference(referencePath, filter, limit, List.of());
    }

    public static String normalize(String val) {
        return RequestFormMasterProductCandidateMatcher.normalize(val);
    }

    public static Map<String, String> loadRowByShohinCode(Path referencePath, String shohinCode)
            throws IOException {
        return PostProcessingProductMasterReferenceCache.rowByCode(referencePath, shohinCode);
    }

    private static String normalizeLength(String val) {
        return RequestFormMasterProductCandidateMatcher.normalizeLengthKeyword(val);
    }

    private static PostProcessingProductMasterIo.SearchHit toHit(
            ProductInfo p, Map<String, Map<String, String>> rowsByCode) {
        if (p == null) {
            return new PostProcessingProductMasterIo.SearchHit("", "", "", "", Map.of());
        }
        String code = normalize(p.getShohinCode());
        Map<String, String> row = rowsByCode.get(code);
        if (row != null && !row.isEmpty()) {
            return new PostProcessingProductMasterIo.SearchHit(
                    row.getOrDefault("商品コード", ""),
                    row.getOrDefault("商品名1", ""),
                    row.getOrDefault("発泡体品番", ""),
                    row.getOrDefault("発泡体品名", ""),
                    Map.copyOf(row));
        }
        Map<String, String> sparse =
                Map.of(
                        "商品コード", p.getShohinCode(),
                        "商品名1", p.getShohinName1(),
                        "発泡体品番", p.getFoamPartNo(),
                        "発泡体品名", p.getFoamName(),
                        "発泡体タイプ", "",
                        "発泡体幅", p.getFoamWidth(),
                        "発泡体長さ", p.getFoamLength(),
                        "発泡体色", p.getFoamColor());
        return new PostProcessingProductMasterIo.SearchHit(
                p.getShohinCode(),
                p.getShohinName1(),
                p.getFoamPartNo(),
                p.getFoamName(),
                sparse);
    }

    private static ProductInfo findByCode(List<ProductInfo> catalog, String normCode) {
        for (ProductInfo p : catalog) {
            if (normCode.equals(normalize(p.getShohinCode()))) {
                return p;
            }
        }
        return null;
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
