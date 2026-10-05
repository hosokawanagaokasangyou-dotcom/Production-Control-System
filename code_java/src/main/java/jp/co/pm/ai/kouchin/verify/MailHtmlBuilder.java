package jp.co.pm.ai.kouchin.verify;

import java.util.ArrayList;
import java.util.List;

/**
 * 報告メールと同内容の Outlook 貼付用 HTML。Excel 全シート変換ではない。
 */
public final class MailHtmlBuilder {

    private MailHtmlBuilder() {}

    public static String build(MailSnapshot kokubu, MailSnapshot konan) {
        List<MailSnapshot> rows = new ArrayList<>();
        if (kokubu != null) {
            rows.add(kokubu);
        }
        if (konan != null) {
            rows.add(konan);
        }
        String ym = firstYm(kokubu, konan);
        String nextYm = firstNext(kokubu, konan);
        String td = "border:1px solid #8FAADC;padding:4px 10px;";
        StringBuilder body = new StringBuilder();
        long totalT1 = 0;
        long totalT2 = 0;
        long totalAdj = 0;
        long totalTg = 0;
        int totalN = 0;
        for (MailSnapshot r : rows) {
            totalT1 += r.total1();
            totalT2 += r.uriage2();
            totalAdj += r.adjustPrev();
            totalTg += r.tougetsu();
            totalN += r.tougetsuCount();
            body.append(tr(td, siteLabel(r), r.total1(), r.uriage2(), r.adjustPrev(), r.tougetsu(), false));
        }
        if (!rows.isEmpty()) {
            body.append(tr(td, "合計", totalT1, totalT2, totalAdj, totalTg, true));
        }
        StringBuilder detail = new StringBuilder();
        for (MailSnapshot r : rows) {
            detail.append("<div>　")
                    .append(esc(r.factoryLabel()))
                    .append("　　")
                    .append(r.tougetsuCount())
                    .append("件（")
                    .append(Fmt.n0(r.tougetsu()))
                    .append("）</div>\n");
        }
        String word = totalTg < 0 ? "不足" : "過剰";
        String highlight = totalTg == 0
                ? totalN + "件の差異（0円）がありました。"
                : totalN + "件の差異（" + Fmt.n0(totalTg) + "円）" + word + "がありました。";
        return "<!DOCTYPE html>\n"
                + "<html><head><meta charset=\"utf-8\"><title>後加工賃支払い明細 "
                + esc(ym)
                + " 検証結果</title></head>\n"
                + "<body style=\"font-family:'Yu Gothic','Meiryo','MS PGothic',sans-serif;font-size:14px;color:#000;line-height:1.7;\">\n"
                + "<div>東レ株式会社</div>\n"
                + "<div>トーレペフ事業部 御中</div>\n"
                + "<br>\n"
                + "<div>いつも大変お世話になっております。</div>\n"
                + "<div>長岡産業の難波です。</div>\n"
                + "<br>\n"
                + "<div>掲題の件、下記にご報告申し上げます。</div>\n"
                + "<br>\n"
                + "<table style=\"border-collapse:collapse;font-size:13px;font-family:'Yu Gothic','Meiryo',sans-serif;\">\n"
                + "<tr>"
                + th("拠点")
                + th("東レ支払いデータ")
                + th("実売上")
                + th("前月調整分")
                + th("当月調整分")
                + th("合計")
                + th("差異 ※")
                + "</tr>\n"
                + "<tr>"
                + tdHead(td, "")
                + tdHead(td, "A")
                + tdHead(td, "①")
                + tdHead(td, "②")
                + tdHead(td, "③")
                + tdHead(td, "B=①+②+③")
                + tdHead(td, "A-B")
                + "</tr>\n"
                + body
                + "</table>\n"
                + "<br>\n"
                + "<div>合計（税抜き） "
                + Fmt.n0(totalT1)
                + "円の御社お支払いデータに対し</div>\n"
                + detail
                + "<div>合計　<span style=\"font-weight:bold;text-decoration:underline;\">"
                + highlight
                + "</span></div>\n"
                + "<br>\n"
                + "<div>添付ファイルをご確認いただき、次月（"
                + esc(nextYm)
                + "）での</div>\n"
                + "<div>ご調整をよろしくお願い申し上げます。</div>\n"
                + "<div>今後とも何卒よろしくお願い申し上げます。</div>\n"
                + "<br>\n"
                + "<div>以上</div>\n"
                + "</body></html>\n";
    }

    private static String th(String label) {
        return "<td style=\"border:1px solid #8FAADC;padding:4px 10px;background:#BDD7EE;font-weight:bold;text-align:center;\">"
                + esc(label)
                + "</td>";
    }

    private static String tdHead(String td, String label) {
        return "<td style=\"" + td + "text-align:center;\">" + esc(label) + "</td>";
    }

    private static String siteLabel(MailSnapshot s) {
        if ("kokubu".equalsIgnoreCase(s.factoryCode())) {
            return "国分（A010）";
        }
        if ("konan".equalsIgnoreCase(s.factoryCode())) {
            return "湖南（A010P）";
        }
        return s.factoryLabel();
    }

    private static String tr(String td, String label, long t1, long t2, long adj, long tg, boolean bold) {
        long b = t2 + adj + tg;
        long diff = t1 - b;
        String w = bold ? "font-weight:bold;" : "";
        return "<tr style=\""
                + w
                + "\">"
                + cell(td, label, "center")
                + cell(td, Fmt.n0(t1), "right")
                + cell(td, Fmt.n0(t2), "right")
                + cell(td, Fmt.n0(adj), "right")
                + cell(td, Fmt.n0(tg), "right")
                + cell(td, Fmt.n0(b), "right")
                + cell(td, Fmt.n0(diff), "right")
                + "</tr>\n";
    }

    private static String cell(String td, String text, String align) {
        return "<td style=\"" + td + "text-align:" + align + ";\">" + esc(text) + "</td>";
    }

    private static String firstYm(MailSnapshot a, MailSnapshot b) {
        if (a != null && a.ymLabel() != null && !a.ymLabel().isBlank()) {
            return a.ymLabel();
        }
        if (b != null && b.ymLabel() != null && !b.ymLabel().isBlank()) {
            return b.ymLabel();
        }
        return "";
    }

    private static String firstNext(MailSnapshot a, MailSnapshot b) {
        if (a != null && a.nextYmLabel() != null && !a.nextYmLabel().isBlank()) {
            return a.nextYmLabel();
        }
        if (b != null && b.nextYmLabel() != null && !b.nextYmLabel().isBlank()) {
            return b.nextYmLabel();
        }
        return "翌月";
    }

    private static String esc(String s) {
        if (s == null) {
            return "";
        }
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;");
    }
}
