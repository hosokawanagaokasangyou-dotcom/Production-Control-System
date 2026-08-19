package jp.co.pm.ai.desktop.reconciliation;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;

/**
 * 依頼書の加工内容・EC面から両面EC/片面ECを判定する。
 *
 * <p>Python {@code planning_core.core.ec_side_classification} と同一ロジック。
 */
public final class EcSideClassification {

    public static final String COLUMN_TITLE = "EC面区分";
    public static final String DOUBLE_SIDED = "両面EC";
    public static final String SINGLE_SIDED = "片面EC";
    public static final String UNKNOWN = "不明";

    private static final List<String> BLANK_EC_MARKERS =
            List.of("-", "―", "－", "—", "nan", "none", "null");

    private EcSideClassification() {}

    public static String classify(String processContent, String ecMen) {
        return classify(processContent, ecMen, true, false);
    }

    public static String classify(String processContent, String ecMen, boolean juchuRowFound) {
        return classify(processContent, ecMen, juchuRowFound, false);
    }

    /**
     * 段階1向け判定。EC面 空かつ依頼書原本参照無しは {@link #UNKNOWN}。
     * EC面 空かつ原本参照ありは {@link #DOUBLE_SIDED}。
     */
    public static String classify(
            String processContent, String ecMen, boolean juchuRowFound, boolean originalRefFound) {
        if (!processContentHasEc(processContent)) {
            return "";
        }
        if (!juchuRowFound) {
            return UNKNOWN;
        }
        if (isBlankEcMen(ecMen)) {
            return originalRefFound ? DOUBLE_SIDED : UNKNOWN;
        }
        if (ecMenIndicatesDoubleSided(ecMen)) {
            return DOUBLE_SIDED;
        }
        if (ecMenIndicatesSingleSided(ecMen)) {
            return SINGLE_SIDED;
        }
        return UNKNOWN;
    }

    static boolean ecMenIndicatesDoubleSided(String ecMen) {
        if (isBlankEcMen(ecMen)) {
            return false;
        }
        String norm = JuchuTransferValueNormalizer.normalizeText(ecMen);
        return norm.contains("両面");
    }

    static boolean ecMenIndicatesSingleSided(String ecMen) {
        if (isBlankEcMen(ecMen) || ecMenIndicatesDoubleSided(ecMen)) {
            return false;
        }
        String norm = JuchuTransferValueNormalizer.normalizeText(ecMen);
        String upper = norm.toUpperCase(Locale.ROOT);
        if (norm.contains("片面")) {
            return true;
        }
        if ("H".equals(upper) || "Q".equals(upper)) {
            return true;
        }
        if (upper.startsWith("H面") || upper.startsWith("Ｈ面")) {
            return true;
        }
        if (upper.startsWith("Q面") || upper.startsWith("Ｑ面")) {
            return true;
        }
        if (norm.contains("ｽﾗｲｽ") || norm.contains("スライス")
                || norm.contains("ｽｷﾝ") || norm.contains("スキン")) {
            return true;
        }
        return !norm.isEmpty();
    }

    /**
     * 配台: 両面EC は EC 工程では2回分。SEC 工程はワンパスで両面EC 可能のため1回分。
     * 国分工場（{@code PM_AI_FACTORY_SITE=KOKUBU}）は1パスで両面EC 可能なため等倍。
     */
    public static int ecDispatchPassCount(String ecSideClass, String processName) {
        return ecDispatchPassCount(ecSideClass, processName, System.getenv("PM_AI_FACTORY_SITE"));
    }

    public static int ecDispatchPassCount(
            String ecSideClass, String processName, String factorySite) {
        if (!DOUBLE_SIDED.equals(ecSideClass != null ? ecSideClass.strip() : "")) {
            return 1;
        }
        if (isKokubuFactory(factorySite)) {
            return 1;
        }
        String proc = normalizeProcessName(processName);
        if ("SEC".equals(proc)) {
            return 1;
        }
        if ("EC".equals(proc)) {
            return 2;
        }
        return 1;
    }

    static boolean isKokubuFactory(String factorySite) {
        if (factorySite == null || factorySite.isBlank()) {
            return false;
        }
        return "KOKUBU".equalsIgnoreCase(factorySite.strip());
    }

    static String normalizeProcessName(String raw) {
        if (raw == null) {
            return "";
        }
        String t = JuchuTransferValueNormalizer.normalizeText(raw.strip());
        return t.replaceAll("[\\s　]+", "");
    }

    public static boolean processContentHasEc(String processContent) {
        for (String tok : parseProcessContentTokens(processContent)) {
            String upper = tok.toUpperCase(Locale.ROOT);
            if ("EC".equals(upper) || upper.startsWith("EC")) {
                return true;
            }
        }
        return false;
    }

    /**
     * 枝番依頼NO の EC 面 lookup 用親キー。
     * 例: {@code W7-22-1} → {@code W7-22}（末尾の {@code -} + 数字 1 セグメントを除く）。
     */
    public static String parentIraiNoLookupKey(String iraiNo) {
        String key = RequestFormOriginalIndexLookup.normalizeIraiNoKey(iraiNo);
        if (key.isEmpty() || !key.contains("-")) {
            return "";
        }
        int lastDash = key.lastIndexOf('-');
        if (lastDash <= 0 || lastDash >= key.length() - 1) {
            return "";
        }
        String tail = key.substring(lastDash + 1);
        if (tail.isEmpty() || !tail.chars().allMatch(Character::isDigit)) {
            return "";
        }
        String head = key.substring(0, lastDash);
        if (!head.contains("-")) {
            return "";
        }
        return head;
    }

    /** 依頼NO 直 lookup → 親依頼NO の順で EC面区分文字列を返す。 */
    public static String resolveEcSideClass(Map<String, String> byKey, String iraiNo) {
        if (byKey == null || byKey.isEmpty()) {
            return "";
        }
        String key = RequestFormOriginalIndexLookup.normalizeIraiNoKey(iraiNo);
        String direct = byKey.get(key);
        if (direct != null && !direct.isBlank()) {
            return direct.strip();
        }
        String parent = parentIraiNoLookupKey(iraiNo);
        if (!parent.isEmpty()) {
            String inherited = byKey.get(parent);
            if (inherited != null && !inherited.isBlank()) {
                return inherited.strip();
            }
        }
        return "";
    }

    static List<String> parseProcessContentTokens(String val) {
        if (JuchuTransferValueNormalizer.isBlank(val)) {
            return List.of();
        }
        String s = JuchuTransferValueNormalizer.normalizeText(val);
        if (s.isEmpty()) {
            return List.of();
        }
        String lower = s.toLowerCase(Locale.ROOT);
        if ("nan".equals(lower) || "none".equals(lower) || "null".equals(lower)) {
            return List.of();
        }
        List<String> out = new ArrayList<>();
        for (String part : s.split(",")) {
            String t = part != null ? part.strip() : "";
            if (!t.isEmpty()) {
                out.add(t);
            }
        }
        return List.copyOf(out);
    }

    private static boolean isBlankEcMen(String ecMen) {
        if (JuchuTransferValueNormalizer.isBlank(ecMen)) {
            return true;
        }
        String norm = JuchuTransferValueNormalizer.normalizeText(ecMen);
        return norm.isEmpty() || BLANK_EC_MARKERS.contains(norm.toLowerCase(Locale.ROOT));
    }
}
