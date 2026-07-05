package jp.co.pm.ai.desktop.reconciliation;

import java.time.LocalDate;
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

    private static final Set<Col> COLOR_COLUMNS = EnumSet.of(Col.IRO_1, Col.IRO);

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
            boolean isMatch =
                    juchuExists && valuesMatch(col, originalValue, juchuValue, juchu);
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
        if (!originalPresent || originalDb == null) {
            return "-";
        }
        if (originalDb.isEmpty()) {
            return "未入力";
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
        for (String segment : raw.split("[\\n/／]+", -1)) {
            String t = segment != null ? segment.strip() : "";
            if (!t.isEmpty() && !parts.contains(t)) {
                parts.add(t);
            }
        }
    }

    private static boolean contractNoValuesMatch(String original, String juchu) {
        List<String> origParts = new ArrayList<>();
        appendContractNoParts(origParts, original);
        List<String> juchuParts = new ArrayList<>();
        appendContractNoParts(juchuParts, juchu);
        if (origParts.isEmpty() && juchuParts.isEmpty()) {
            return true;
        }
        if (origParts.size() != juchuParts.size()) {
            return false;
        }
        for (int i = 0; i < origParts.size(); i++) {
            if (!normalizeContractNoToken(origParts.get(i))
                    .equals(normalizeContractNoToken(juchuParts.get(i)))) {
                return false;
            }
        }
        return true;
    }

    /** 目次シート vs 依頼シートの契約Ｎｏ比較。 */
    static boolean contractNoStringsEquivalent(String a, String b) {
        return contractNoValuesMatch(a, b);
    }

    /** {@code P000075564} と {@code P75564} を同一視。 */
    private static String normalizeContractNoToken(String token) {
        String t = JuchuTransferValueNormalizer.normalizeText(token);
        if (t.matches("(?i)P\\d+")) {
            return "P" + Long.parseLong(t.substring(1));
        }
        return t;
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
        if (valuesMatch(Col.KEIYAKU_NO, originalContract, juchuContract, juchu)) {
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

    private static boolean valuesMatch(
            Col col, String original, String juchu, Map<String, String> juchuDb) {
        if (col == Col.USER) {
            return userValuesMatch(original, juchu);
        }
        if (col == Col.KEIYAKU_NO) {
            return contractNoValuesMatch(original, juchu);
        }
        if (col == Col.SEIHIN) {
            return productSpecValuesMatch(original, juchu);
        }
        if (col == Col.EC_MEN) {
            return ecMenValuesMatch(original, juchu, juchuDb);
        }
        if (COLOR_COLUMNS.contains(col)) {
            return colorValuesMatch(original, juchu);
        }
        if (col == Col.KAKO_NAIYO) {
            return normalizeProcessContent(original).equals(normalizeProcessContent(juchu));
        }
        if (DATE_COLUMNS.contains(col)) {
            return datesMatch(original, juchu);
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

    private static boolean datesMatch(String original, String juchu) {
        List<String> origLines = splitDateLines(original);
        List<String> juchuLines = splitDateLines(juchu);
        if (origLines.isEmpty() && juchuLines.isEmpty()) {
            return true;
        }
        if (origLines.isEmpty() || juchuLines.isEmpty()) {
            return false;
        }
        if (origLines.size() == juchuLines.size()) {
            for (int i = 0; i < origLines.size(); i++) {
                if (!singleDateMatch(origLines.get(i), juchuLines.get(i))) {
                    return false;
                }
            }
            return true;
        }
        // 原反複数行で原本のみ M/D が複数行・受注が1行（同一日）など
        if (juchuLines.size() == 1) {
            String juchuDate = juchuLines.get(0);
            return origLines.stream().allMatch(o -> singleDateMatch(o, juchuDate));
        }
        if (origLines.size() == 1) {
            String origDate = origLines.get(0);
            return juchuLines.stream().allMatch(j -> singleDateMatch(origDate, j));
        }
        return datesMatchSameUniqueResolvedDate(origLines, juchuLines);
    }

    /** 行数は異なるが、解決後の日付がいずれも同一なら一致。 */
    private static boolean datesMatchSameUniqueResolvedDate(
            List<String> origLines, List<String> juchuLines) {
        LocalDate juchuRef =
                juchuLines.stream()
                        .map(JuchuTransferValueNormalizer::parseLocalDate)
                        .filter(d -> d != null)
                        .findFirst()
                        .orElse(LocalDate.now());
        LocalDate origRef =
                origLines.stream()
                        .map(JuchuTransferValueNormalizer::parseLocalDate)
                        .filter(d -> d != null)
                        .findFirst()
                        .orElse(juchuRef);
        List<LocalDate> origResolved = resolveDateLines(origLines, juchuRef);
        List<LocalDate> juchuResolved = resolveDateLines(juchuLines, origRef);
        if (origResolved.contains(null) || juchuResolved.contains(null)) {
            return false;
        }
        if (origResolved.stream().distinct().count() != 1
                || juchuResolved.stream().distinct().count() != 1) {
            return false;
        }
        return origResolved.get(0).equals(juchuResolved.get(0));
    }

    private static List<LocalDate> resolveDateLines(List<String> lines, LocalDate yearReference) {
        List<LocalDate> resolved = new ArrayList<>();
        for (String line : lines) {
            resolved.add(JuchuTransferValueNormalizer.parseLocalDate(line, yearReference));
        }
        return resolved;
    }

    private static List<String> splitDateLines(String val) {
        if (JuchuTransferValueNormalizer.isBlank(val)) {
            return List.of();
        }
        List<String> lines = new ArrayList<>();
        for (String line : val.split("\\n", -1)) {
            String t = line != null ? line.strip() : "";
            if (!t.isEmpty()) {
                lines.add(t);
            }
        }
        return List.copyOf(lines);
    }

    private static boolean singleDateMatch(String original, String juchu) {
        LocalDate juchuFull = JuchuTransferValueNormalizer.parseLocalDate(juchu);
        LocalDate originalFull = JuchuTransferValueNormalizer.parseLocalDate(original);
        LocalDate origResolved =
                JuchuTransferValueNormalizer.parseLocalDate(
                        original, juchuFull != null ? juchuFull : LocalDate.now());
        LocalDate juchuResolved =
                JuchuTransferValueNormalizer.parseLocalDate(
                        juchu, originalFull != null ? originalFull : LocalDate.now());
        if (origResolved != null && juchuResolved != null) {
            return origResolved.equals(juchuResolved);
        }
        return JuchuTransferValueNormalizer.normalizeDateVal(original)
                .equals(JuchuTransferValueNormalizer.normalizeDateVal(juchu));
    }

    private static boolean userValuesMatch(String original, String juchu) {
        String ru = JuchuTransferValueNormalizer.normalizeText(original);
        String dbu = JuchuTransferValueNormalizer.normalizeText(juchu);
        return ru.equals(dbu) || ru.contains(dbu) || dbu.contains(ru);
    }

    /** 受注側の品番プレフィックス（例: {@code 30020-A05W-870-870X97}）を除いて製品 spec を比較。 */
    private static boolean productSpecValuesMatch(String original, String juchu) {
        return normalizeProductSpec(original).equals(normalizeProductSpec(juchu));
    }

    private static String normalizeProductSpec(String val) {
        if (val == null || val.isBlank()) {
            return "";
        }
        String normalized = JuchuTransferValueNormalizer.normalizeText(val);
        // 受注の品番プレフィックス（例: 30020-）を除去
        normalized = normalized.replaceFirst("^\\d+-", "");
        // TPI 原本の幅重複（A05W-870-870X97）を受注形式（A05W-870X97）へ
        normalized =
                normalized.replaceFirst(
                        "^(A\\d{2}W|R\\d{2}W)-(\\d+)-\\2X(\\d+)$", "$1-$2X$3");
        return normalized;
    }

    /** 受注 EC 欄が空でも、加工内容に EC 面情報があれば一致扱い。 */
    private static boolean ecMenValuesMatch(
            String original, String juchu, Map<String, String> juchuDb) {
        if (!JuchuTransferValueNormalizer.isBlank(juchu)) {
            return JuchuTransferValueNormalizer.normalizeText(original)
                    .equals(JuchuTransferValueNormalizer.normalizeText(juchu));
        }
        if (JuchuTransferValueNormalizer.isBlank(original)) {
            return true;
        }
        String kako = valueForKey(juchuDb, Col.KAKO_NAIYO.dbKey());
        if (JuchuTransferValueNormalizer.isBlank(kako)) {
            return false;
        }
        String normKako = JuchuTransferValueNormalizer.normalizeText(kako);
        String normOrig = JuchuTransferValueNormalizer.normalizeText(original);
        if (normKako.contains(normOrig)) {
            return true;
        }
        if (normOrig.contains("片面") && normKako.contains("片面") && normKako.contains("EC")) {
            return true;
        }
        if (normOrig.contains("両面") && normKako.contains("両面") && normKako.contains("EC")) {
            return true;
        }
        return false;
    }

    /** {@code LG} と {@code ライトグレー} 等の略称・正式名を同一視。 */
    private static boolean colorValuesMatch(String original, String juchu) {
        return normalizeColorKey(original).equals(normalizeColorKey(juchu));
    }

    private static String normalizeColorKey(String val) {
        if (val == null || val.isBlank()) {
            return "";
        }
        String normalized = JuchuTransferValueNormalizer.normalizeText(val);
        if ("LG".equals(normalized)
                || (normalized.contains("ライト") && normalized.contains("グレ"))) {
            return "LIGHTGRAY";
        }
        return normalized;
    }

    private static String normalizeProcessContent(String val) {
        return JuchuTransferValueNormalizer.normalizeText(val).replace(",", "").replace("、", "");
    }

    private static String nullToEmpty(String val) {
        return val != null ? val : "";
    }
}
