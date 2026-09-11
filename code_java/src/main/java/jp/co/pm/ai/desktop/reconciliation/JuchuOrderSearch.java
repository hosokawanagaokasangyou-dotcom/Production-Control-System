package jp.co.pm.ai.desktop.reconciliation;

import java.text.Collator;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Collection;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.function.Function;

public final class JuchuOrderSearch {

    private JuchuOrderSearch() {}

    public static LocalDate defaultDeliveryFrom(LocalDate today) {
        LocalDate t = today != null ? today : LocalDate.now();
        return t.minusMonths(6);
    }

    public static LocalDate defaultDeliveryTo(LocalDate today) {
        return today != null ? today : LocalDate.now();
    }

    public static List<OrderRecord> filter(
            Collection<OrderRecord> records, JuchuOrderSearchCriteria criteria) {
        Objects.requireNonNull(criteria, "criteria");
        var validationError = criteria.validationError();
        if (validationError.isPresent()) {
            throw new IllegalArgumentException(validationError.get());
        }
        if (records == null) {
            return List.of();
        }
        return records.stream()
                .filter(record -> matches(record, criteria, "", ""))
                .toList();
    }

    public static List<OrderRecord> filter(
            Collection<OrderRecord> records,
            JuchuOrderSearchCriteria criteria,
            Function<OrderRecord, String> extraMachineHaystack,
            Function<OrderRecord, String> extraProcessHaystack) {
        Objects.requireNonNull(criteria, "criteria");
        var validationError = criteria.validationError();
        if (validationError.isPresent()) {
            throw new IllegalArgumentException(validationError.get());
        }
        if (records == null) {
            return List.of();
        }
        Function<OrderRecord, String> machines =
                extraMachineHaystack != null ? extraMachineHaystack : r -> "";
        Function<OrderRecord, String> processes =
                extraProcessHaystack != null ? extraProcessHaystack : r -> "";
        return records.stream()
                .filter(
                        record ->
                                matches(
                                        record,
                                        criteria,
                                        machines.apply(record),
                                        processes.apply(record)))
                .toList();
    }

    public static boolean matches(OrderRecord record, JuchuOrderSearchCriteria criteria) {
        return matches(record, criteria, "", "");
    }

    public static boolean matches(
            OrderRecord record,
            JuchuOrderSearchCriteria criteria,
            String extraMachineHaystack,
            String extraProcessHaystack) {
        if (record == null || criteria == null) {
            return false;
        }
        if (!deliveryInRange(record, criteria.from(), criteria.to())) {
            return false;
        }
        return keywordMatches(record, criteria, extraMachineHaystack, extraProcessHaystack);
    }

    public static String displayRawMaterial(Map<String, String> dbValues) {
        if (dbValues == null) {
            return "";
        }
        return firstNonBlank(dbValues.get("原反"), dbValues.get("品名1"), dbValues.get("原反品名"));
    }

    private static boolean deliveryInRange(OrderRecord record, LocalDate from, LocalDate to) {
        Map<String, String> db = record.getDbValues();
        if (db == null) {
            return false;
        }
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

    public static String displayMachine(Map<String, String> dbValues, String extraHaystack) {
        return firstNonBlank(
                dbValues != null ? dbValues.get("機械名") : null,
                dbValues != null ? dbValues.get("機械") : null,
                extraHaystack);
    }

    public static String displayProcess(Map<String, String> dbValues, String extraHaystack) {
        return firstNonBlank(
                dbValues != null ? dbValues.get("工程名") : null,
                dbValues != null ? dbValues.get("加工内容") : null,
                extraHaystack);
    }

    public static List<String> productCandidates(Collection<OrderRecord> records) {
        return distinctSorted(collectField(records, r -> dbField(r, "製品")));
    }

    public static List<String> rawMaterialCandidates(Collection<OrderRecord> records) {
        return distinctSorted(collectField(records, JuchuOrderSearch::displayRawMaterialFromRecord));
    }

    public static List<String> machineCandidates(
            Collection<OrderRecord> records, Collection<String> extraNames) {
        Set<String> values = collectField(records, r -> displayMachine(r.getDbValues(), ""));
        addAllCandidates(values, extraNames);
        return distinctSorted(values);
    }

    public static List<String> processCandidates(
            Collection<OrderRecord> records, Collection<String> extraNames) {
        Set<String> values = collectField(records, r -> displayProcess(r.getDbValues(), ""));
        addAllCandidates(values, extraNames);
        return distinctSorted(values);
    }

    static List<String> distinctSorted(Collection<String> values) {
        Set<String> unique = new LinkedHashSet<>();
        addAllCandidates(unique, values);
        List<String> out = new ArrayList<>(unique);
        Collator ja = Collator.getInstance(Locale.JAPAN);
        ja.setStrength(Collator.PRIMARY);
        out.sort(ja);
        return List.copyOf(out);
    }

    private static boolean keywordMatches(
            OrderRecord record,
            JuchuOrderSearchCriteria criteria,
            String extraMachineHaystack,
            String extraProcessHaystack) {
        Map<String, String> db = record.getDbValues();
        if (db == null) {
            return false;
        }
        String productKeyword = normalizedKeyword(criteria.productKeyword());
        String rawMaterialKeyword = normalizedKeyword(criteria.rawMaterialKeyword());
        String machineKeyword = normalizedKeyword(criteria.machineKeyword());
        String processKeyword = normalizedKeyword(criteria.processKeyword());

        boolean productMatch =
                !productKeyword.isEmpty()
                        && containsNormalized(db.get("製品"), productKeyword);
        boolean rawMaterialMatch =
                !rawMaterialKeyword.isEmpty()
                        && (containsNormalized(db.get("原反"), rawMaterialKeyword)
                                || containsNormalized(db.get("品名1"), rawMaterialKeyword)
                                || containsNormalized(db.get("原反品名"), rawMaterialKeyword));
        boolean machineMatch =
                !machineKeyword.isEmpty()
                        && (containsNormalized(db.get("機械名"), machineKeyword)
                                || containsNormalized(db.get("機械"), machineKeyword)
                                || containsNormalized(extraMachineHaystack, machineKeyword));
        boolean processMatch =
                !processKeyword.isEmpty()
                        && (containsNormalized(db.get("工程名"), processKeyword)
                                || containsNormalized(db.get("加工内容"), processKeyword)
                                || containsNormalized(extraProcessHaystack, processKeyword));
        boolean productOrRawOk =
                (productKeyword.isEmpty() && rawMaterialKeyword.isEmpty())
                        || productMatch
                        || rawMaterialMatch;
        boolean machineOk = machineKeyword.isEmpty() || machineMatch;
        boolean processOk = processKeyword.isEmpty() || processMatch;
        return productOrRawOk && machineOk && processOk;
    }

    private static String firstNonBlank(String... values) {
        if (values == null) {
            return "";
        }
        for (String value : values) {
            if (value != null && !value.strip().isEmpty()) {
                return value.strip();
            }
        }
        return "";
    }

    private static String displayRawMaterialFromRecord(OrderRecord record) {
        return displayRawMaterial(record != null ? record.getDbValues() : null);
    }

    private static String dbField(OrderRecord record, String key) {
        if (record == null || record.getDbValues() == null) {
            return "";
        }
        return nullToEmpty(record.getDbValues().get(key));
    }

    private static Set<String> collectField(
            Collection<OrderRecord> records, Function<OrderRecord, String> extractor) {
        Set<String> out = new LinkedHashSet<>();
        if (records == null || extractor == null) {
            return out;
        }
        for (OrderRecord record : records) {
            addCandidates(out, extractor.apply(record));
        }
        return out;
    }

    private static void addAllCandidates(Set<String> out, Collection<String> values) {
        if (out == null || values == null) {
            return;
        }
        for (String value : values) {
            addCandidates(out, value);
        }
    }

    private static void addCandidates(Set<String> out, String raw) {
        if (out == null || raw == null || raw.isBlank()) {
            return;
        }
        for (String line : raw.split("\\R")) {
            if (line == null) {
                continue;
            }
            String v = line.strip();
            if (!v.isEmpty()) {
                out.add(v);
            }
        }
    }

    private static String nullToEmpty(String value) {
        return value != null ? value : "";
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
