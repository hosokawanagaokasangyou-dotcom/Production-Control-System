package jp.co.pm.ai.desktop.io.actuals;

import java.io.PrintWriter;
import java.io.StringWriter;

/** ダッシュボード読込失敗時の例外整形（UI 非依存）。 */
public final class DashboardLoadErrorFormatter {

    /** 原因連鎖をたどる上限。循環・過度に深いラップで表示が壊れないようにする。 */
    static final int MAX_CAUSE_DEPTH = 6;

    private DashboardLoadErrorFormatter() {}

    /** 例外クラス名とメッセージを「原因:」で連結した1〜数行の要約。 */
    public static String formatDetail(Throwable ex) {
        if (ex == null) {
            return "原因不明";
        }
        StringBuilder sb = new StringBuilder();
        Throwable cur = ex;
        int depth = 0;
        while (cur != null && depth < MAX_CAUSE_DEPTH) {
            if (depth > 0) {
                sb.append("\n原因: ");
            }
            sb.append(cur.getClass().getSimpleName());
            String msg = cur.getMessage();
            if (msg != null && !msg.isBlank()) {
                sb.append(": ").append(msg.strip());
            }
            cur = cur.getCause();
            depth++;
        }
        return sb.toString();
    }

    /** 要約の先頭1行のみ。サブバーなど1行しか使えない箇所向け。 */
    public static String formatShortDetail(Throwable ex) {
        String detail = formatDetail(ex);
        int nl = detail.indexOf('\n');
        return nl >= 0 ? detail.substring(0, nl) : detail;
    }

    public static String formatStackTrace(Throwable ex) {
        if (ex == null) {
            return "";
        }
        StringWriter sw = new StringWriter();
        ex.printStackTrace(new PrintWriter(sw));
        return sw.toString().strip();
    }
}
