package jp.co.pm.ai.planning.stage2.core;

import java.io.BufferedReader;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.Normalizer;
import java.util.Collections;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;

/**
 * Python {@code planning_core._core} のロール単位長さテーブル（{@code code/使用原反,ロール単位の長さ.txt} /
 * {@code code/製品名,ロール単位の長さ.txt}）を読み、照会キーを正規化して引く。
 */
public final class Stage2RollUnitLengthTables {

    private static final String USED_RAW_FILENAME = "使用原反,ロール単位の長さ.txt";
    private static final String PRODUCT_FILENAME = "製品名,ロール単位の長さ.txt";

    private final Map<String, Double> byUsedRaw;
    private final Map<String, Double> byProduct;

    private Stage2RollUnitLengthTables(Map<String, Double> byUsedRaw, Map<String, Double> byProduct) {
        this.byUsedRaw = byUsedRaw;
        this.byProduct = byProduct;
    }

    public static Stage2RollUnitLengthTables empty() {
        return new Stage2RollUnitLengthTables(Map.of(), Map.of());
    }

    public static Stage2RollUnitLengthTables load(Path repoRoot) throws IOException {
        if (repoRoot == null || !Files.isDirectory(repoRoot)) {
            return empty();
        }
        Path code = repoRoot.resolve("code");
        return new Stage2RollUnitLengthTables(
                readTable(code.resolve(USED_RAW_FILENAME)),
                readTable(code.resolve(PRODUCT_FILENAME)));
    }

    public Optional<Double> lookupByUsedRaw(String usedRaw) {
        String k = normalizeKey(usedRaw);
        if (k.isEmpty()) {
            return Optional.empty();
        }
        Double v = byUsedRaw.get(k);
        return v != null && v > 0 ? Optional.of(v) : Optional.empty();
    }

    public Optional<Double> lookupByProductName(String productName) {
        String k = normalizeKey(productName);
        if (k.isEmpty()) {
            return Optional.empty();
        }
        Double v = byProduct.get(k);
        return v != null && v > 0 ? Optional.of(v) : Optional.empty();
    }

    public static String normalizeKey(String val) {
        if (val == null) {
            return "";
        }
        String s = Normalizer.normalize(val.strip(), Normalizer.Form.NFKC);
        StringBuilder b = new StringBuilder(s.length());
        for (int i = 0; i < s.length(); i++) {
            char ch = s.charAt(i);
            if (!Character.isWhitespace(ch)) {
                b.append(ch);
            }
        }
        return b.toString();
    }

    private static Map<String, Double> readTable(Path path) throws IOException {
        if (!Files.isRegularFile(path)) {
            return Collections.emptyMap();
        }
        Map<String, Double> out = new HashMap<>();
        try (BufferedReader br = Files.newBufferedReader(path, StandardCharsets.UTF_8)) {
            String line;
            boolean firstNonEmpty = true;
            while ((line = br.readLine()) != null) {
                line = line.strip();
                if (line.isEmpty()) {
                    continue;
                }
                if (firstNonEmpty) {
                    firstNonEmpty = false;
                    if (line.contains("ロール単位")) {
                        continue;
                    }
                }
                int c = line.lastIndexOf(',');
                if (c <= 0 || c >= line.length() - 1) {
                    continue;
                }
                String rawKey = line.substring(0, c).strip();
                String rawM = line.substring(c + 1).strip();
                String key = normalizeKey(rawKey);
                if (key.isEmpty()) {
                    continue;
                }
                double m = parseDoubleLenient(rawM);
                if (m <= 0) {
                    continue;
                }
                out.putIfAbsent(key, m);
            }
        }
        return out;
    }

    private static double parseDoubleLenient(String rawM) {
        try {
            return Double.parseDouble(rawM.replace(",", "."));
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }

    /** Python {@code _infer_roll_unit_m_from_product_name_dimensions_only} に概ね相当。 */
    public static double inferFromProductDimensions(String productName, double fallbackUnit) {
        if (productName == null || productName.isBlank()) {
            return fallbackUnit;
        }
        String s = normalizeProductDimSeparators(productName);
        java.util.regex.Pattern pair =
                java.util.regex.Pattern.compile("(\\d{2,6})\\s*[xX]\\s*(\\d{2,6})");
        java.util.regex.Matcher mp = pair.matcher(s);
        int lastB = -1;
        while (mp.find()) {
            try {
                lastB = Integer.parseInt(mp.group(2));
            } catch (NumberFormatException ignored) {
                lastB = -1;
            }
        }
        if (lastB > 0) {
            return lastB;
        }
        java.util.regex.Pattern tail = java.util.regex.Pattern.compile("[xX]\\s*(\\d{2,6})");
        java.util.regex.Matcher mt = tail.matcher(s);
        int v = -1;
        while (mt.find()) {
            try {
                v = Integer.parseInt(mt.group(1));
            } catch (NumberFormatException ignored) {
                v = -1;
            }
        }
        if (v > 0) {
            return v;
        }
        return fallbackUnit > 0 ? fallbackUnit : 100.0;
    }

    private static String normalizeProductDimSeparators(String s) {
        String t = Normalizer.normalize(s, Normalizer.Form.NFKC);
        for (String ch :
                new String[] {
                    "\u00d7",
                    "\u2715",
                    "\u2716",
                    "\u2a2f",
                    "\u2a09",
                    "\uff38",
                    "\uff58",
                    "\u0425",
                    "\u0445",
                    "\u03a7",
                    "\u03c7",
                }) {
            t = t.replace(ch, "x");
        }
        return t;
    }

    public static double parseFloatSafe(String v, double defaultVal) {
        if (v == null) {
            return defaultVal;
        }
        String s = v.strip();
        if (s.isEmpty()
                || s.equalsIgnoreCase("nan")
                || s.equalsIgnoreCase("none")
                || s.equals("-")
                || s.equals("—")
                || s.equals("―")) {
            return defaultVal;
        }
        try {
            return Double.parseDouble(s.replace(",", "."));
        } catch (NumberFormatException e) {
            return defaultVal;
        }
    }

    public static Optional<Double> optionalUnprocessedCell(String cell) {
        if (cell == null) {
            return Optional.empty();
        }
        String s = Normalizer.normalize(cell.strip(), Normalizer.Form.NFKC);
        if (s.isEmpty()
                || s.equalsIgnoreCase("nan")
                || s.equalsIgnoreCase("none")
                || s.equals("-")
                || s.equals("—")
                || s.equals("―")) {
            return Optional.empty();
        }
        try {
            return Optional.of(Double.parseDouble(s.replace(",", ".")));
        } catch (NumberFormatException e) {
            return Optional.empty();
        }
    }

    public static String formatMetersPlain(double x) {
        if (Double.isNaN(x) || Double.isInfinite(x)) {
            return "";
        }
        if (Math.abs(x - Math.rint(x)) < 1e-9) {
            return String.format(Locale.ROOT, "%.0f", x);
        }
        return String.format(Locale.ROOT, "%s", x);
    }

    public static String formatPercentPlain(double ratio01) {
        if (Double.isNaN(ratio01) || Double.isInfinite(ratio01)) {
            return "";
        }
        double pct = 100.0 * ratio01;
        if (Math.abs(pct - Math.rint(pct)) < 1e-6) {
            return String.format(Locale.ROOT, "%.0f%%", pct);
        }
        return String.format(Locale.ROOT, "%.2f%%", pct);
    }
}
