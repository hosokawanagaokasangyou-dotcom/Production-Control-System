package jp.co.pm.ai.kouchin.verify;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/** 検証結果ブック内の金額セル番地。 */
public final class AmountCellRef {

    private AmountCellRef() {}

    public static String address(String sheet, int zeroBasedColumn, int row) {
        String name = sheet == null ? "" : sheet.replace("'", "''");
        return "'" + name + "'!" + column(zeroBasedColumn) + row;
    }

    /** 0始まりの列番号を A, B, …, AA にする。 */
    public static String column(int zeroBasedColumn) {
        int n = zeroBasedColumn + 1;
        StringBuilder sb = new StringBuilder();
        while (n > 0) {
            int rem = (n - 1) % 26;
            sb.append((char) ('A' + rem));
            n = (n - 1) / 26;
        }
        return sb.reverse().toString();
    }

    public static Map<String, List<String>> freeze(Map<String, List<String>> cells) {
        Map<String, List<String>> copy = new LinkedHashMap<>();
        if (cells == null) {
            return Map.of();
        }
        for (Map.Entry<String, List<String>> e : cells.entrySet()) {
            if (e.getKey() == null || e.getValue() == null || e.getValue().isEmpty()) {
                continue;
            }
            copy.put(e.getKey(), List.copyOf(e.getValue()));
        }
        return Map.copyOf(copy);
    }
}
