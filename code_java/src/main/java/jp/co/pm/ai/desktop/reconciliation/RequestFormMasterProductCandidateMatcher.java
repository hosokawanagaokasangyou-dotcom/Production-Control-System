package jp.co.pm.ai.desktop.reconciliation;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;

/**
 * アラジン統合マスタから依頼書フォーム用の商品候補を、一致度スコア順に列挙する。
 * 全キーワード AND 一致ではなく、部分一致・数値近似も含めて複数候補を返す。
 */
final class RequestFormMasterProductCandidateMatcher {

    private static final int MIN_SCORE_TO_LIST = 20;

    private RequestFormMasterProductCandidateMatcher() {}

    static List<String> buildRankedCandidateLabels(
            List<ProductInfo> catalog,
            String kwItem,
            String kwPart,
            String kwType,
            String kwLength,
            String kwHinmei,
            int limit) {
        if (catalog == null || catalog.isEmpty() || limit <= 0) {
            return List.of();
        }
        String item = normalize(kwItem);
        String part = normalize(kwPart);
        String type = normalize(kwType);
        String length = normalizeLengthKeyword(kwLength);
        String hinmei = normalize(kwHinmei);

        boolean anyKeyword =
                !item.isEmpty() || !part.isEmpty() || !type.isEmpty() || !length.isEmpty() || !hinmei.isEmpty();
        if (!anyKeyword) {
            List<String> labels = new ArrayList<>(Math.min(limit, catalog.size()));
            for (int i = 0; i < catalog.size() && labels.size() < limit; i++) {
                labels.add(formatCandidateLabel(catalog.get(i)));
            }
            return labels;
        }

        List<ScoredProduct> scored = new ArrayList<>();
        for (ProductInfo product : catalog) {
            int score = scoreProduct(product, item, part, type, length, hinmei);
            if (score >= MIN_SCORE_TO_LIST) {
                scored.add(new ScoredProduct(score, product));
            }
        }
        scored.sort(
                Comparator.comparingInt(ScoredProduct::score)
                        .reversed()
                        .thenComparing(s -> s.product().getShohinCode()));

        List<String> labels = new ArrayList<>(Math.min(limit, scored.size()));
        for (int i = 0; i < scored.size() && labels.size() < limit; i++) {
            labels.add(formatCandidateLabel(scored.get(i).product()));
        }
        return labels;
    }

    private static int scoreProduct(
            ProductInfo p, String item, String part, String type, String length, String hinmei) {
        int total = 0;
        int active = 0;
        int matched = 0;

        if (!item.isEmpty()) {
            active++;
            int s =
                    bestFieldScore(
                            item, p.getShohinCode(), p.getSeihinCode(), p.getShohinName1(), p.getShohinName2());
            if (s > 0) {
                matched++;
                total += s;
            }
        }
        if (!part.isEmpty()) {
            active++;
            int s = bestFieldScore(part, p.getFoamPartNo(), p.getShohinCode());
            if (s > 0) {
                matched++;
                total += s;
            }
        }
        if (!type.isEmpty()) {
            active++;
            int s = bestFieldScore(type, p.getShohinName1(), p.getFoamName());
            if (s > 0) {
                matched++;
                total += s;
            }
        }
        if (!length.isEmpty()) {
            active++;
            String pLength = normalizeLengthKeyword(p.getFoamLength());
            int s = bestFieldScore(length, pLength);
            if (s > 0) {
                matched++;
                total += s;
            }
        }
        if (!hinmei.isEmpty()) {
            active++;
            int s =
                    bestFieldScore(
                            hinmei,
                            p.getFoamName(),
                            p.getShohinCode(),
                            p.getSeihinCode(),
                            p.getFoamPartNo());
            if (s > 0) {
                matched++;
                total += s;
            }
        }

        if (active > 0 && matched == active) {
            total += 35;
        } else if (matched > 0 && matched < active) {
            total += 10;
        }
        return total;
    }

    private static int bestFieldScore(String keyword, String... fields) {
        int max = 0;
        if (fields != null) {
            for (String field : fields) {
                max = Math.max(max, fieldScore(keyword, field));
            }
        }
        return max;
    }

    private static int fieldScore(String keyword, String field) {
        if (keyword.isEmpty() || field == null || field.isBlank()) {
            return 0;
        }
        String k = normalize(field);
        if (k.isEmpty()) {
            return 0;
        }
        if (k.equals(keyword)) {
            return 100;
        }
        if (k.contains(keyword)) {
            return 78;
        }
        if (keyword.length() >= 3 && keyword.contains(k)) {
            return 55;
        }
        return Math.max(digitSimilarityScore(keyword, k), prefixSimilarityScore(keyword, k));
    }

    private static int digitSimilarityScore(String keyword, String field) {
        String kd = digitsOnly(keyword);
        String fd = digitsOnly(field);
        if (kd.isEmpty() || fd.isEmpty()) {
            return 0;
        }
        if (fd.equals(kd)) {
            return 92;
        }
        if (fd.contains(kd)) {
            return 72;
        }
        if (kd.contains(fd) && fd.length() >= 3) {
            return 58;
        }
        int dist = levenshtein(kd, fd);
        int maxLen = Math.max(kd.length(), fd.length());
        if (maxLen >= 3 && dist <= 2) {
            return 58 - dist * 14;
        }
        if (maxLen >= 4 && dist <= maxLen / 2) {
            return 40 - dist * 6;
        }
        return 0;
    }

    private static int prefixSimilarityScore(String keyword, String field) {
        int common = 0;
        int n = Math.min(keyword.length(), field.length());
        for (int i = 0; i < n; i++) {
            if (keyword.charAt(i) == field.charAt(i)) {
                common++;
            } else {
                break;
            }
        }
        if (common >= 4) {
            return common * 10;
        }
        if (common >= 2) {
            return common * 6;
        }
        return 0;
    }

    private static int levenshtein(String a, String b) {
        int[][] dp = new int[a.length() + 1][b.length() + 1];
        for (int i = 0; i <= a.length(); i++) {
            dp[i][0] = i;
        }
        for (int j = 0; j <= b.length(); j++) {
            dp[0][j] = j;
        }
        for (int i = 1; i <= a.length(); i++) {
            for (int j = 1; j <= b.length(); j++) {
                int cost = a.charAt(i - 1) == b.charAt(j - 1) ? 0 : 1;
                dp[i][j] =
                        Math.min(
                                Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1),
                                dp[i - 1][j - 1] + cost);
            }
        }
        return dp[a.length()][b.length()];
    }

    private static String digitsOnly(String text) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < text.length(); i++) {
            char ch = text.charAt(i);
            if (Character.isDigit(ch)) {
                sb.append(ch);
            }
        }
        return sb.toString();
    }

    private static String normalizeLengthKeyword(String val) {
        return normalize(val).replaceAll("\\.0$", "");
    }

    static String normalize(String val) {
        if (val == null) {
            return "";
        }
        String text = val.strip();
        text = java.text.Normalizer.normalize(text, java.text.Normalizer.Form.NFKC);
        text = text.replaceAll("\\s+", "");
        text = text.replace("－", "-").replace("ー", "-").replace("―", "-").replace("‐", "-");
        return text.toUpperCase(Locale.ROOT);
    }

    static String formatCandidateLabel(ProductInfo p) {
        String pLength = p.getFoamLength() != null ? p.getFoamLength().replaceAll("\\.0$", "") : "";
        String pWidth = p.getFoamWidth() != null ? p.getFoamWidth().replaceAll("\\.0$", "") : "";
        String dims = (pWidth.isEmpty() ? "?" : pWidth) + "×" + (pLength.isEmpty() ? "?" : pLength);
        return p.getShohinCode()
                + " | "
                + p.getFoamPartNo()
                + " | "
                + p.getFoamName()
                + " | "
                + dims;
    }

    private record ScoredProduct(int score, ProductInfo product) {}
}
