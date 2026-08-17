package jp.co.pm.ai.desktop.io;

import java.util.List;

import jp.co.pm.ai.desktop.ui.ClipboardTableSupport;

/**
 * 同一化チェック差異の表・クリップボード用変換。
 */
public final class AladdinEntryIdentityCheckDiffTable {

    public static final List<String> HEADERS =
            List.of("機械", "依頼NO", "工程", "日付", "シス計", "加工計画");

    private AladdinEntryIdentityCheckDiffTable() {}

    public static String toCsv(List<AladdinEntryDispatchPlanIdentityCheck.Diff> diffs) {
        return toDelimited(diffs, ',');
    }

    public static String toTsv(List<AladdinEntryDispatchPlanIdentityCheck.Diff> diffs) {
        return toDelimited(diffs, '\t');
    }

    public static String toHtmlTable(List<AladdinEntryDispatchPlanIdentityCheck.Diff> diffs) {
        if (diffs == null || diffs.isEmpty()) {
            return "";
        }
        StringBuilder sb = new StringBuilder();
        sb.append(
                "<table border=\"1\" cellspacing=\"0\" cellpadding=\"4\""
                        + " style=\"border-collapse:collapse;font-family:'Meiryo UI',sans-serif;font-size:11pt;\">");
        sb.append("<thead><tr>");
        for (String header : HEADERS) {
            sb.append("<th style=\"background:#D9E1F2;padding:4px 8px;text-align:left;\">")
                    .append(ClipboardTableSupport.escapeHtml(header))
                    .append("</th>");
        }
        sb.append("</tr></thead><tbody>");
        for (AladdinEntryDispatchPlanIdentityCheck.Diff d : diffs) {
            if (d == null) {
                continue;
            }
            sb.append("<tr>");
            for (String cell : cells(d)) {
                sb.append("<td style=\"padding:4px 8px;\">")
                        .append(ClipboardTableSupport.escapeHtml(cell))
                        .append("</td>");
            }
            sb.append("</tr>");
        }
        sb.append("</tbody></table>");
        return sb.toString();
    }

    public static List<String> cells(AladdinEntryDispatchPlanIdentityCheck.Diff d) {
        if (d == null) {
            return List.of();
        }
        return List.of(
                nz(d.machineName()),
                nz(d.taskId()),
                nz(d.processName()),
                d.date() != null ? d.date().toString() : "",
                formatQty(d.systemQty()),
                formatQty(d.planQty()));
    }

    private static String toDelimited(
            List<AladdinEntryDispatchPlanIdentityCheck.Diff> diffs, char delimiter) {
        if (diffs == null || diffs.isEmpty()) {
            return "";
        }
        StringBuilder sb = new StringBuilder();
        appendRow(sb, HEADERS, delimiter);
        for (AladdinEntryDispatchPlanIdentityCheck.Diff d : diffs) {
            if (d == null) {
                continue;
            }
            sb.append('\n');
            appendRow(sb, cells(d), delimiter);
        }
        return sb.toString();
    }

    private static void appendRow(StringBuilder sb, List<String> cells, char delimiter) {
        for (int i = 0; i < cells.size(); i++) {
            if (i > 0) {
                sb.append(delimiter);
            }
            appendCell(sb, cells.get(i), delimiter);
        }
    }

    private static void appendCell(StringBuilder sb, String value, char delimiter) {
        String text = value != null ? value : "";
        if (text.indexOf(delimiter) >= 0
                || text.indexOf('"') >= 0
                || text.indexOf('\n') >= 0
                || text.indexOf('\r') >= 0) {
            sb.append('"').append(text.replace("\"", "\"\"")).append('"');
        } else {
            sb.append(text);
        }
    }

    public static String formatQty(double qty) {
        if (Math.abs(qty - Math.rint(qty)) < 1e-9) {
            return Long.toString(Math.round(qty));
        }
        return Double.toString(qty);
    }

    private static String nz(String s) {
        return s != null ? s : "";
    }
}
