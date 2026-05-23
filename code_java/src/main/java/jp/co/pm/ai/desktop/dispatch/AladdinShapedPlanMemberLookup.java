package jp.co.pm.ai.desktop.dispatch;

import java.util.List;
import java.util.regex.Pattern;

/**
 * アラジン加工計画（shaped JSON / 表）から {@code 機械名×依頼NO×工程} に紐づく担当者名を解決する。
 *
 * <p>加工実績明細に {@code メンバー名} が無い／空のとき、{@code 担当OP_指定} 等の静的列を参照する。
 */
public final class AladdinShapedPlanMemberLookup {

    private static final Pattern ALADDIN_DATE_COL = Pattern.compile("\\d{4}/\\d{2}/\\d{2}");

    private static final String COL_MK_NAME = "機械名";
    private static final String COL_TID = "依頼NO";
    private static final String COL_PROCESS = "工程名";

    /** Python {@code PLAN_COL_PREFERRED_OP} および配台表の表記揺れ。 */
    private static final String[] MEMBER_COLUMNS = {
        "担当OP_指定", "担当OP指定", "メンバー名"
    };

    private AladdinShapedPlanMemberLookup() {}

    /**
     * 担当者名を返す。見つからなければ空文字。
     *
     * @param dateStr {@code yyyy/MM/dd}（日付列に担当名テキストが入るレイアウト向けフォールバック）
     */
    public static String lookup(
            List<String> headers,
            List<List<String>> rows,
            String machine,
            String requestNo,
            String processRaw,
            String dateStr) {
        if (headers == null || rows == null || rows.isEmpty()) {
            return "";
        }
        int mkIdx = colIdx(headers, COL_MK_NAME);
        int tidIdx = colIdx(headers, COL_TID);
        int procIdx = colIdx(headers, COL_PROCESS);
        if (mkIdx < 0 || tidIdx < 0) {
            return "";
        }
        String mkNorm = normalizeEquipmentMatchKey(machine);
        String tid = requestNo != null ? requestNo.strip() : "";
        if (mkNorm.isEmpty() || tid.isEmpty()) {
            return "";
        }
        String procKey = AladdinShapedPlanQtyLookup.normalizeProcessNameForRuleMatch(processRaw);
        Integer dateColIdx = dateColumnIndex(headers, dateStr);

        String loose = "";
        for (List<String> row : rows) {
            if (!mkNorm.equals(normalizeEquipmentMatchKey(cellAt(row, mkIdx)))) {
                continue;
            }
            if (!tid.equals(cellAt(row, tidIdx).strip())) {
                continue;
            }
            String rowProcKey =
                    procIdx >= 0
                            ? AladdinShapedPlanQtyLookup.normalizeProcessNameForRuleMatch(
                                    cellAt(row, procIdx))
                            : "";
            boolean processExact =
                    procKey.isEmpty()
                            || rowProcKey.isEmpty()
                            || procKey.equals(rowProcKey);
            if (!processExact) {
                continue;
            }
            String fromStatic = firstNonBlankMemberColumn(headers, row);
            if (!fromStatic.isEmpty()) {
                return fromStatic;
            }
            if (dateColIdx != null) {
                String cell = cellAt(row, dateColIdx).strip();
                if (!cell.isEmpty() && !looksLikeNumericQty(cell)) {
                    return cell;
                }
            }
            if (loose.isEmpty()) {
                loose = firstNonBlankMemberColumn(headers, row);
            }
        }
        return loose;
    }

    private static String firstNonBlankMemberColumn(List<String> headers, List<String> row) {
        for (String col : MEMBER_COLUMNS) {
            int idx = colIdx(headers, col);
            if (idx >= 0) {
                String v = cellAt(row, idx).strip();
                if (!v.isEmpty()) {
                    return v;
                }
            }
        }
        return "";
    }

    private static Integer dateColumnIndex(List<String> headers, String dateStr) {
        if (dateStr == null || dateStr.isBlank()) {
            return null;
        }
        String key = AladdinShapedPlanQtyLookup.normaliseDateStr(dateStr);
        if (key == null) {
            key = dateStr.strip();
        }
        for (int i = 0; i < headers.size(); i++) {
            String h = headers.get(i);
            if (h == null || !ALADDIN_DATE_COL.matcher(h).matches()) {
                continue;
            }
            String hk = AladdinShapedPlanQtyLookup.normaliseDateStr(h);
            if (hk == null) {
                hk = h;
            }
            if (key.equals(hk)) {
                return i;
            }
        }
        return null;
    }

    private static boolean looksLikeNumericQty(String raw) {
        if (raw == null || raw.isBlank()) {
            return false;
        }
        try {
            Double.parseDouble(raw.strip().replace(",", ""));
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    private static String normalizeEquipmentMatchKey(String val) {
        if (val == null || val.isBlank()) {
            return "";
        }
        String t = java.text.Normalizer.normalize(val, java.text.Normalizer.Form.NFKC);
        t = t.replace('\u00a0', ' ').replace('\u3000', ' ');
        t = t.replaceAll("[\u200b\u200c\u200d\ufeff]", "");
        return t.replaceAll("\\s+", " ").strip();
    }

    private static int colIdx(List<String> headers, String title) {
        for (int i = 0; i < headers.size(); i++) {
            if (title.equals(headers.get(i))) {
                return i;
            }
        }
        return -1;
    }

    private static String cellAt(List<String> row, int idx) {
        return (idx >= 0 && idx < row.size() && row.get(idx) != null) ? row.get(idx) : "";
    }
}
