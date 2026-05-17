package jp.co.pm.ai.desktop.io;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import java.util.regex.Pattern;

import jp.co.pm.ai.desktop.config.SummaryAiDispatchExportPrefs;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchSchema;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo.TabularSheet;

/**
 * サマリ Excel 出力向け: 日付列の判定と、非日付列のみの並べ替え。
 */
public final class SummaryAiDispatchExportColumnSupport {

    private static final Pattern MAIN_COMPARE_DATE_HDR =
            Pattern.compile("(\\d{4})\u5e74(\\d{1,2})\u6708(\\d{1,2})\u65e5\\([\u6708\u706b\u6c34\u6728\u91d1\u571f\u65e5]\\)");

    private static final Pattern SLASH_DATE_HDR = Pattern.compile("\\d{4}/\\d{2}/\\d{2}");

    private SummaryAiDispatchExportColumnSupport() {}

    public static boolean isDateColumnHeader(SummaryAiDispatchExportPrefs.SheetKey sheet, String header) {
        if (header == null || header.isBlank()) {
            return false;
        }
        return switch (sheet) {
            case MAIN_COMPARE -> MAIN_COMPARE_DATE_HDR.matcher(header).matches();
            case DISPATCH ->
                    ResultDispatchSchema.isDateColumn(header)
                            || SLASH_DATE_HDR.matcher(header).matches();
            case ALADDIN, ACTUALS -> SLASH_DATE_HDR.matcher(header).matches();
        };
    }

    static TabularSheet applySheetLayout(
            TabularSheet data, SummaryAiDispatchExportPrefs.SheetKey sheetKey, List<String> savedOrder) {
        if (data == null) {
            return new TabularSheet(List.of(), List.of());
        }
        List<String> headers = data.headers() != null ? new ArrayList<>(data.headers()) : new ArrayList<>();
        List<List<String>> rows = data.rows() != null ? data.rows() : List.of();
        if (headers.isEmpty()) {
            return data;
        }
        List<Integer> perm = buildExportPermutation(headers, sheetKey, savedOrder);
        if (isIdentityPerm(perm)) {
            return data;
        }
        List<String> newHeaders = new ArrayList<>();
        for (int i : perm) {
            newHeaders.add(headers.get(i));
        }
        List<List<String>> newRows = new ArrayList<>();
        for (List<String> row : rows) {
            List<String> line = new ArrayList<>(newHeaders.size());
            for (int i : perm) {
                String v =
                        row != null && i < row.size() && row.get(i) != null ? row.get(i) : "";
                line.add(v);
            }
            newRows.add(line);
        }
        return new TabularSheet(newHeaders, newRows);
    }

    /** 非日付列のみの見出し（保存・UI 表示用）。 */
    static List<String> nonDateHeaders(List<String> headers, SummaryAiDispatchExportPrefs.SheetKey sheet) {
        List<String> out = new ArrayList<>();
        if (headers == null) {
            return out;
        }
        for (String h : headers) {
            if (!isDateColumnHeader(sheet, h)) {
                out.add(h);
            }
        }
        return out;
    }

    private static List<Integer> buildExportPermutation(
            List<String> headers, SummaryAiDispatchExportPrefs.SheetKey sheet, List<String> savedOrder) {
        List<String> leading = new ArrayList<>();
        List<String> dates = new ArrayList<>();
        for (String h : headers) {
            if (isDateColumnHeader(sheet, h)) {
                dates.add(h);
            } else {
                leading.add(h);
            }
        }
        List<String> orderedLeading = reorderNonDate(leading, savedOrder);
        List<String> targetOrder = new ArrayList<>(orderedLeading);
        targetOrder.addAll(dates);
        return buildPermutation(headers, targetOrder);
    }

    private static List<String> reorderNonDate(List<String> current, List<String> savedOrder) {
        if (savedOrder == null || savedOrder.isEmpty()) {
            return current;
        }
        List<String> out = new ArrayList<>();
        Set<String> used = new HashSet<>();
        for (String title : savedOrder) {
            for (String h : current) {
                if (Objects.equals(h, title) && used.add(h)) {
                    out.add(h);
                    break;
                }
            }
        }
        for (String h : current) {
            if (used.add(h)) {
                out.add(h);
            }
        }
        return out;
    }

    private static List<Integer> buildPermutation(List<String> fileHeaders, List<String> targetOrder) {
        List<Integer> perm = new ArrayList<>();
        Set<Integer> used = new HashSet<>();
        for (String title : targetOrder) {
            int idx = findNextUnusedMatching(fileHeaders, title, used);
            if (idx >= 0) {
                perm.add(idx);
                used.add(idx);
            }
        }
        for (int i = 0; i < fileHeaders.size(); i++) {
            if (!used.contains(i)) {
                perm.add(i);
                used.add(i);
            }
        }
        return perm;
    }

    private static int findNextUnusedMatching(List<String> headers, String title, Set<Integer> used) {
        for (int i = 0; i < headers.size(); i++) {
            if (used.contains(i)) {
                continue;
            }
            if (Objects.equals(headers.get(i), title)) {
                return i;
            }
        }
        return -1;
    }

    private static boolean isIdentityPerm(List<Integer> perm) {
        for (int i = 0; i < perm.size(); i++) {
            if (perm.get(i) != i) {
                return false;
            }
        }
        return true;
    }
}
