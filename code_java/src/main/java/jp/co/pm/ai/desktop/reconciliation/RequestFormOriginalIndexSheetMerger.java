package jp.co.pm.ai.desktop.reconciliation;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/** 目次シートの値を依頼シート rawMap へ優先反映し、相違メタデータを付与する。 */
final class RequestFormOriginalIndexSheetMerger {

    static final String META_INDEX_APPLIED = "_indexSheetApplied";
    static final String META_INDEX_CONFLICTS = "_indexSheetConflicts";

    private static final String KEY_INPUT_DATE = "投入日";
    private static final String KEY_DELIVERY_ANSWER = "納期回答";
    private static final String KEY_CONTRACT_NO = "契約Ｎｏ";

    private RequestFormOriginalIndexSheetMerger() {}

    static void applyIndexOverrides(
            Map<String, String> rawMap,
            RequestFormOriginalIndexSheetReader.IndexEntry indexEntry) {
        if (rawMap == null || indexEntry == null) {
            return;
        }
        rawMap.put(META_INDEX_APPLIED, "true");
        List<String> conflicts = new ArrayList<>();

        // 目次値で上書きする前に依頼シート単体の投入日を保持（原反投入日4ソース照合用）。
        String sheetInputDate = nullToEmpty(rawMap.get(KEY_INPUT_DATE)).strip();
        if (!sheetInputDate.isEmpty()) {
            rawMap.put(RequestFormOriginalIndexSheetMeta.KEY_SHEET_INPUT_DATE, sheetInputDate);
        }

        applyField(
                rawMap,
                conflicts,
                KEY_INPUT_DATE,
                "投入日",
                indexEntry.inputDate(),
                RequestFormOriginalIndexSheetMerger::inputDatesEquivalent);

        applyField(
                rawMap,
                conflicts,
                KEY_DELIVERY_ANSWER,
                "納期回答",
                indexEntry.deliveryDate(),
                RequestFormOriginalIndexSheetMerger::datesEquivalent);

        applyField(
                rawMap,
                conflicts,
                KEY_CONTRACT_NO,
                "契約Ｎｏ",
                indexEntry.contractNo(),
                JuchuTransferCoverageCheck::contractNoStringsEquivalent);

        if (conflicts.isEmpty()) {
            rawMap.remove(META_INDEX_CONFLICTS);
        } else {
            rawMap.put(META_INDEX_CONFLICTS, formatConflictBanner(conflicts));
        }

        RequestFormOriginalIndexSheetMeta.writeIndexMetaToRawMap(rawMap, indexEntry);
    }

    private static void applyField(
            Map<String, String> rawMap,
            List<String> conflicts,
            String rawKey,
            String label,
            String indexValue,
            BiEquivalent comparator) {
        if (JuchuTransferValueNormalizer.isBlank(indexValue)) {
            return;
        }
        String sheetValue = nullToEmpty(rawMap.get(rawKey));
        String indexText = indexValue.strip();
        if (!comparator.equivalent(sheetValue, indexText)) {
            conflicts.add(formatConflictLine(label, sheetValue, indexText));
        }
        rawMap.put(rawKey, indexText);
    }

    private static String formatConflictLine(String label, String sheetValue, String indexValue) {
        String sheetDisplay = sheetValue.isBlank() ? "（空）" : sheetValue;
        return "・" + label + ": シート " + sheetDisplay + " → 目次 " + indexValue;
    }

    private static String formatConflictBanner(List<String> lines) {
        StringBuilder sb = new StringBuilder();
        sb.append("【目次シート優先】依頼シートと目次の記載が異なります。");
        sb.append("転記・照合には目次の値を使用しています。");
        sb.append(System.lineSeparator());
        for (String line : lines) {
            sb.append(line);
            sb.append(System.lineSeparator());
        }
        return sb.toString().strip();
    }

    private static boolean datesEquivalent(String sheetValue, String indexValue) {
        if (JuchuTransferValueNormalizer.isBlank(indexValue)) {
            return true;
        }
        if (JuchuTransferValueNormalizer.isBlank(sheetValue)) {
            return false;
        }
        String normIndex = JuchuTransferValueNormalizer.normalizeDateVal(indexValue);
        for (String part : sheetValue.split("\\n", -1)) {
            if (JuchuTransferValueNormalizer.isBlank(part)) {
                continue;
            }
            if (normIndex.equals(JuchuTransferValueNormalizer.normalizeDateVal(part))) {
                return true;
            }
        }
        return false;
    }

    private static boolean inputDatesEquivalent(String sheetValue, String indexValue) {
        return datesEquivalent(sheetValue, indexValue);
    }

    private static String nullToEmpty(String value) {
        return value != null ? value : "";
    }

    @FunctionalInterface
    private interface BiEquivalent {
        boolean equivalent(String sheetValue, String indexValue);
    }
}
