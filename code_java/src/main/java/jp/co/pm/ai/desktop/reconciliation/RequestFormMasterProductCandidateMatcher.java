package jp.co.pm.ai.desktop.reconciliation;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;

/**
 * アラジン統合マスタから依頼書フォーム用の商品候補を、一致度スコア順に列挙する。
 * 品名（フォームの品名＝マスタのフォーム銘 foamName）は正規化後の完全一致のみ。近似一致はしない。
 * <p>項目ごとの重み（{@link #WEIGHT_SCALE} 倍スケール）: 商品コード 1.5、品番 2、タイプ 3、長さ・品名 1。
 */
final class RequestFormMasterProductCandidateMatcher {

    private static final int MIN_SCORE_TO_LIST = 20;

    /** 重みの小数表現用（1.0 = 10）。 */
    private static final int WEIGHT_SCALE = 10;
    private static final int WEIGHT_ITEM = 15;
    private static final int WEIGHT_PART = 20;
    private static final int WEIGHT_TYPE = 30;
    private static final int WEIGHT_LENGTH = 10;
    private static final int WEIGHT_HINMEI = 10;

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
            if (!hinmei.isEmpty() && !hinmeiMatchesProduct(hinmei, product)) {
                continue;
            }
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
                total += applyFieldWeight(s, WEIGHT_ITEM);
            }
        }
        if (!part.isEmpty()) {
            active++;
            int s = bestFieldScore(part, p.getFoamPartNo(), p.getShohinCode());
            if (s > 0) {
                matched++;
                total += applyFieldWeight(s, WEIGHT_PART);
            }
        }
        if (!type.isEmpty()) {
            active++;
            int s =
                    bestFieldScore(
                            type,
                            formatTypeForLabel(p.getShohinName1()),
                            p.getShohinName1(),
                            p.getFoamName());
            if (s > 0) {
                matched++;
                total += applyFieldWeight(s, WEIGHT_TYPE);
            }
        }
        if (!length.isEmpty()) {
            active++;
            String pLength = normalizeLengthKeyword(p.getFoamLength());
            int s = bestFieldScore(length, pLength);
            if (s > 0) {
                matched++;
                total += applyFieldWeight(s, WEIGHT_LENGTH);
            }
        }
        if (!hinmei.isEmpty()) {
            active++;
            int s = hinmeiFieldScore(hinmei, p);
            if (s > 0) {
                matched++;
                total += applyFieldWeight(s, WEIGHT_HINMEI);
            }
        }

        if (active > 0 && matched == active) {
            total += 35;
        } else if (matched > 0 && matched < active) {
            total += 10;
        }
        return total;
    }

    private static int applyFieldWeight(int rawScore, int weightScaled) {
        if (rawScore <= 0 || weightScaled <= 0) {
            return 0;
        }
        return (rawScore * weightScaled) / WEIGHT_SCALE;
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
        String kwd = normalize(keyword);
        String k = normalize(field);
        if (kwd.isEmpty() || k.isEmpty()) {
            return 0;
        }
        if (k.equals(kwd)) {
            return 100;
        }
        if (k.contains(kwd)) {
            return 78;
        }
        if (kwd.length() >= 3 && kwd.contains(k)) {
            return 55;
        }
        return 0;
    }

    /** 品名は foamName と正規化後完全一致のみ（6783 と 6798 は別候補）。 */
    private static boolean hinmeiMatchesProduct(String hinmei, ProductInfo product) {
        return hinmei.equals(normalize(product.getFoamName()));
    }

    private static int hinmeiFieldScore(String hinmei, ProductInfo product) {
        return hinmeiMatchesProduct(hinmei, product) ? 100 : 0;
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
        String color = formatFoamColorForLabel(p.getFoamColor());
        String kako = p.getKakoNaiyo();
        if (kako == null || kako.isBlank()) {
            kako = "?";
        }
        return p.getShohinCode()
                + " | "
                + p.getFoamPartNo()
                + " | "
                + p.getFoamName()
                + " | "
                + formatTypeForLabel(p.getShohinName1())
                + " | "
                + dims
                + " | "
                + color
                + " | "
                + kako;
    }

    /** 候補表示用タイプ（フォーム転記と同じ {@code shohinName1} の解釈）。 */
    static String formatTypeForLabel(String shohinName1) {
        if (shohinName1 == null || shohinName1.isBlank()) {
            return "?";
        }
        String[] nameParts = shohinName1.split("-");
        if (nameParts.length >= 2 && !nameParts[1].isBlank()) {
            return nameParts[1].strip();
        }
        return shohinName1.strip();
    }

    /** 候補表示用。空は {@code ?}、数値 0 は {@code -}。 */
    static String formatFoamColorForLabel(String foamColor) {
        if (foamColor == null || foamColor.isBlank()) {
            return "?";
        }
        String text = foamColor.strip().replaceAll("\\.0$", "");
        if (isZeroFoamColor(text)) {
            return "-";
        }
        return foamColor.strip();
    }

    private static boolean isZeroFoamColor(String text) {
        if (text.isEmpty()) {
            return false;
        }
        if ("0".equals(text)) {
            return true;
        }
        try {
            return Double.parseDouble(text) == 0.0d;
        } catch (NumberFormatException ex) {
            return false;
        }
    }

    private record ScoredProduct(int score, ProductInfo product) {}
}
