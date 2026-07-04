package jp.co.pm.ai.desktop.reconciliation;

import java.util.ArrayList;
import java.util.EnumSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import jp.co.pm.ai.desktop.reconciliation.JuchuSheetColumnLayout.Col;

/**
 * 依頼書原本の転記対象項目が受注ファイルに正しく書き込まれているかを列単位で照合する。
 */
public final class JuchuTransferCoverageCheck {

    private static final Set<Col> DATE_COLUMNS =
            EnumSet.of(Col.NYURYOKU_BI, Col.TONYU_BI, Col.KIBO_NOKI, Col.CHOSEI_NOKI);

    private static final Set<Col> NUMERIC_COLUMNS =
            EnumSet.of(Col.SURYO_1, Col.SURYO, Col.WARISU, Col.GENPAN_ROLL_SU, Col.KAKOCHIN);

    private JuchuTransferCoverageCheck() {}

    public record ColumnCheck(
            Col col, String formLabel, String originalValue, String juchuValue, boolean matched) {}

    public record CoverageResult(
            boolean juchuRowExists,
            int totalWithOriginalValue,
            int matchedCount,
            double ratePercent,
            List<ColumnCheck> details) {

        public int mismatchCount() {
            return totalWithOriginalValue - matchedCount;
        }

        public String rateDisplay() {
            if (!juchuRowExists) {
                return "0% (0/" + totalWithOriginalValue + ")";
            }
            int pct = totalWithOriginalValue == 0 ? 100 : (int) Math.round(ratePercent);
            return pct + "% (" + matchedCount + "/" + totalWithOriginalValue + ")";
        }
    }

    public static CoverageResult compare(
            Map<String, String> originalDb,
            Map<String, String> juchuDb,
            JuchuHeaderAliasRegistry registry,
            String juchuFileAbsolutePath) {
        boolean juchuExists = juchuDb != null && !juchuDb.isEmpty();
        Map<String, String> orig = originalDb != null ? originalDb : Map.of();
        Map<String, String> juchu = juchuDb != null ? juchuDb : Map.of();

        List<ColumnCheck> details = new ArrayList<>();
        int total = 0;
        int matched = 0;

        for (Col col : JuchuSheetColumnLayout.transferColumns()) {
            if (registry != null
                    && juchuFileAbsolutePath != null
                    && registry.isExcludedFromTransfer(juchuFileAbsolutePath, col)) {
                continue;
            }
            String dbKey = col.dbKey();
            String originalValue = valueForKey(orig, dbKey);
            if (JuchuTransferValueNormalizer.isBlank(originalValue)) {
                continue;
            }
            total++;
            String juchuValue = valueForKey(juchu, dbKey);
            boolean isMatch = juchuExists && valuesMatch(col, originalValue, juchuValue);
            if (isMatch) {
                matched++;
            }
            details.add(
                    new ColumnCheck(
                            col, col.formItemDescription(), originalValue, juchuValue, isMatch));
        }

        double rate = total == 0 ? 100.0 : (100.0 * matched / total);
        if (!juchuExists) {
            matched = 0;
            rate = 0.0;
            details =
                    details.stream()
                            .map(
                                    d ->
                                            new ColumnCheck(
                                                    d.col(),
                                                    d.formLabel(),
                                                    d.originalValue(),
                                                    d.juchuValue(),
                                                    false))
                            .toList();
        }
        return new CoverageResult(juchuExists, total, matched, rate, List.copyOf(details));
    }

    /**
     * 受注ファイルの契約Ｎｏ表示。値そのものを返す（複数行・複数セル行は {@code /} 連結）。未入力は {@code 未入力}。
     */
    public static String formatJuchuContractNoDisplay(
            Map<String, String> juchuDb, boolean juchuRowExists) {
        if (!juchuRowExists || juchuDb == null || juchuDb.isEmpty()) {
            return "未入力";
        }
        List<String> parts = collectContractNoParts(juchuDb);
        if (parts.isEmpty()) {
            return "未入力";
        }
        return String.join("/", parts);
    }

    /**
     * 依頼書原本の契約Ｎｏ表示。値そのもの（複数行は {@code /} 連結）。原本なしは {@code -}、空欄は {@code 未入力}。
     */
    public static String formatOriginalContractNoDisplay(
            Map<String, String> originalDb, boolean originalPresent) {
        if (!originalPresent || originalDb == null || originalDb.isEmpty()) {
            return "-";
        }
        List<String> parts = collectContractNoParts(originalDb);
        if (parts.isEmpty()) {
            return "未入力";
        }
        return String.join("/", parts);
    }

    /** 契約Ｎｏ文字列（改行区切り可）から非空パーツを順序保持で収集。 */
    static List<String> collectContractNoParts(Map<String, String> map) {
        String raw = contractNoRaw(map);
        List<String> parts = new ArrayList<>();
        appendContractNoParts(parts, raw);
        return List.copyOf(parts);
    }

    /** 改行区切りの契約Ｎｏを {@code target} にマージ（重複は除く）。 */
    static void mergeContractNoValues(Map<String, String> target, Map<String, String> additional) {
        if (target == null || additional == null) {
            return;
        }
        List<String> parts = new ArrayList<>(collectContractNoParts(target));
        appendContractNoParts(parts, contractNoRaw(additional));
        if (parts.isEmpty()) {
            target.remove(Col.KEIYAKU_NO.dbKey());
            target.remove("契約No");
            target.remove("契約NO");
            return;
        }
        target.put(Col.KEIYAKU_NO.dbKey(), String.join("\n", parts));
    }

    private static void appendContractNoParts(List<String> parts, String raw) {
        if (JuchuTransferValueNormalizer.isBlank(raw)) {
            return;
        }
        for (String line : raw.split("\\n", -1)) {
            String t = line != null ? line.strip() : "";
            if (!t.isEmpty() && !parts.contains(t)) {
                parts.add(t);
            }
        }
    }

    private static String contractNoRaw(Map<String, String> map) {
        if (map == null || map.isEmpty()) {
            return "";
        }
        String v = valueForKey(map, Col.KEIYAKU_NO.dbKey());
        if (!JuchuTransferValueNormalizer.isBlank(v)) {
            return v;
        }
        for (String key : List.of("契約No", "契約NO")) {
            v = nullToEmpty(map.get(key));
            if (!JuchuTransferValueNormalizer.isBlank(v)) {
                return v;
            }
        }
        return "";
    }

    /**
     * @deprecated 一覧の契約NO列は {@link #formatJuchuContractNoDisplay} を使用。
     */
    @Deprecated
    public static String contractNoJuchuStatus(
            Map<String, String> originalDb, Map<String, String> juchuDb, boolean juchuRowExists) {
        Map<String, String> orig = originalDb != null ? originalDb : Map.of();
        Map<String, String> juchu = juchuDb != null ? juchuDb : Map.of();
        String originalContract = valueForKey(orig, Col.KEIYAKU_NO.dbKey());
        if (JuchuTransferValueNormalizer.isBlank(originalContract)) {
            return "-";
        }
        if (!juchuRowExists) {
            return "なし(未登録)";
        }
        String juchuContract = valueForKey(juchu, Col.KEIYAKU_NO.dbKey());
        if (JuchuTransferValueNormalizer.isBlank(juchuContract)) {
            return "なし";
        }
        if (valuesMatch(Col.KEIYAKU_NO, originalContract, juchuContract)) {
            return "あり";
        }
        return "相違";
    }

    private static String valueForKey(Map<String, String> map, String dbKey) {
        if (map.containsKey(dbKey)) {
            return nullToEmpty(map.get(dbKey));
        }
        return "";
    }

    private static boolean valuesMatch(Col col, String original, String juchu) {
        if (col == Col.USER) {
            return userValuesMatch(original, juchu);
        }
        if (col == Col.KAKO_NAIYO) {
            return normalizeProcessContent(original).equals(normalizeProcessContent(juchu));
        }
        if (DATE_COLUMNS.contains(col)) {
            return JuchuTransferValueNormalizer.normalizeDateVal(original)
                    .equals(JuchuTransferValueNormalizer.normalizeDateVal(juchu));
        }
        if (NUMERIC_COLUMNS.contains(col)) {
            return Math.abs(
                            JuchuTransferValueNormalizer.normalizeNumeric(original)
                                    - JuchuTransferValueNormalizer.normalizeNumeric(juchu))
                    < 1e-9;
        }
        return JuchuTransferValueNormalizer.normalizeText(original)
                .equals(JuchuTransferValueNormalizer.normalizeText(juchu));
    }

    private static boolean userValuesMatch(String original, String juchu) {
        String ru = JuchuTransferValueNormalizer.normalizeText(original);
        String dbu = JuchuTransferValueNormalizer.normalizeText(juchu);
        return ru.equals(dbu) || ru.contains(dbu) || dbu.contains(ru);
    }

    private static String normalizeProcessContent(String val) {
        return JuchuTransferValueNormalizer.normalizeText(val).replace(",", "").replace("、", "");
    }

    private static String nullToEmpty(String val) {
        return val != null ? val : "";
    }
}
