package jp.co.pm.ai.desktop.reconciliation;

import java.time.LocalDate;
import java.util.List;
import java.util.Map;

public final class JuchuOrderSearch {

    private JuchuOrderSearch() {}

    public static List<OrderRecord> filter(
            List<OrderRecord> records, JuchuOrderSearchCriteria criteria) {
        return records.stream().filter(record -> matches(record, criteria)).toList();
    }

    public static boolean matches(OrderRecord record, JuchuOrderSearchCriteria criteria) {
        if (!deliveryInRange(record, criteria.from(), criteria.to())) {
            return false;
        }
        return keywordMatches(record, criteria);
    }

    public static String displayRawMaterial(Map<String, String> dbValues) {
        if (dbValues == null) {
            return "";
        }
        String hinmei1 = dbValues.get("品名1");
        if (hinmei1 != null && !hinmei1.strip().isEmpty()) {
            return hinmei1.strip();
        }
        String rawName = dbValues.get("原反品名");
        return rawName != null ? rawName.strip() : "";
    }

    private static boolean deliveryInRange(OrderRecord record, LocalDate from, LocalDate to) {
        Map<String, String> db = record.getDbValues();
        return dateInRange(db.get("希望納期"), from, to)
                || dateInRange(db.get("調整納期"), from, to);
    }

    private static boolean dateInRange(String value, LocalDate from, LocalDate to) {
        LocalDate parsed = JuchuTransferValueNormalizer.parseLocalDate(value);
        if (parsed == null) {
            return false;
        }
        return !parsed.isBefore(from) && !parsed.isAfter(to);
    }

    private static boolean keywordMatches(OrderRecord record, JuchuOrderSearchCriteria criteria) {
        Map<String, String> db = record.getDbValues();
        String productKeyword = normalizedKeyword(criteria.productKeyword());
        String rawMaterialKeyword = normalizedKeyword(criteria.rawMaterialKeyword());

        boolean productMatch =
                !productKeyword.isEmpty()
                        && containsNormalized(db.get("製品"), productKeyword);
        boolean rawMaterialMatch =
                !rawMaterialKeyword.isEmpty()
                        && (containsNormalized(db.get("品名1"), rawMaterialKeyword)
                                || containsNormalized(db.get("原反品名"), rawMaterialKeyword));
        return productMatch || rawMaterialMatch;
    }

    private static String normalizedKeyword(String keyword) {
        if (keyword == null || keyword.strip().isEmpty()) {
            return "";
        }
        return JuchuTransferValueNormalizer.normalizeText(keyword);
    }

    private static boolean containsNormalized(String fieldValue, String normalizedKeyword) {
        if (normalizedKeyword.isEmpty()) {
            return false;
        }
        return JuchuTransferValueNormalizer.normalizeText(fieldValue).contains(normalizedKeyword);
    }
}
