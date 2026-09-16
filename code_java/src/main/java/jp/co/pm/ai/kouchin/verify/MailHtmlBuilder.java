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
        String th = "border:1px solid #8FAADC;padding:4px 10px;background:#BDD7EE;font-weight:bold;text-align:center;";
        StringBuilder body = new StringBuilder();
        long totalT1 = 0;
        long totalT2 = 0;
        long totalAdj = 0;
        long totalTg = 0;
        int totalN = 0;
        for (MailSnapshot r : rows) {
            long adj = -r.adjustPrev();
            long tg = -r.tougetsu();
            totalT1 += r.total1();
            totalT2 += r.uriage2();
            totalAdj += adj;
            totalTg += tg;
            totalN += r.tougetsuCount();
            body.append(tr(td, r.factoryLabel(), r.total1(), r.uriage2(), adj, tg, false));
        }
        if (!rows.isEmpty()) {
            body.append(tr(td, "合計", totalT1, totalT2, totalAdj, totalTg, true));
        }
        StringBuilder detail = new StringBuilder();
        for (MailSnapshot r : rows) {
            detail.append("<div>　")
                    .append(esc(r.factoryLabel()))
                    .append("　　　 ")
                    .append(r.tougetsuCount())
                    .append("件（")
                    .append(Fmt.n0(-r.tougetsu()))
                    .append("）</div>\n");
        }
        String highlight = totalN + "件の差異（" + Fmt.n0(Math.abs(totalTg)) + "円）過不足";
        return "<!DOCTYPE html>\n"
                + "<html><head><meta charset=\"utf-8\"><title>後加工賃支払い明細 "
                + esc(ym)
                + " 検証結果</title></head>\n"
                + "<body style=\"font-family:'Yu Gothic','Meiryo','MS PGothic',sans-serif;font-size:14px;color:#000;line-height:1.7;\">\n"
                + "<div>東レ株式会社</div>\n"
                + "<div>トーレペフ事業部 御中</div>\n"
                + "<br>\n"
                + "<div>いつも大変お世話になっております。</div>\n"
                + "<br>\n"
                + "<div>掲題の件、下記にご報告申し上げます。</div>\n"
                + "<br>\n"
                + "<table style=\"border-collapse:collapse;font-size:13px;font-family:'Yu Gothic','Meiryo',sans-serif;\">\n"
                + "<tr>"
                + "<td style=\""
                + th
                + "\">拠点</td>"
                + "<td style=\""
                + th
                + "\">東レ支払いデータ</td>"
                + "<td style=\""
                + th
                + "\">実売上</td>"
                + "<td style=\""
                + th
                + "\">前月調整分</td>"
                + "<td style=\""
                + th
                + "\">当月調整分</td>"
                + "</tr>\n"
                + body
                + "</table>\n"
                + "<br>\n"
                + "<div>合計（税抜き）　"
                + Fmt.n0(totalT1)
                + "円の御社お支払いデータに対し</div>\n"
                + detail
                + "<div>合計　<span style=\"font-weight:bold;text-decoration:underline;\">"
                + highlight
                + "</span>がありました。</div>\n"
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

    private static String tr(String td, String label, long t1, long t2, long adj, long tg, boolean bold) {
        String w = bold ? "font-weight:bold;" : "";
        return "<tr style=\""
                + w
                + "\">"
                + "<td style=\""
                + td
                + "text-align:center;\">"
                + esc(label)
                + "</td>"
                + "<td style=\""
                + td
                + "text-align:right;\">"
                + Fmt.n0(t1)
                + "</td>"
                + "<td style=\""
                + td
                + "text-align:right;\">"
                + Fmt.n0(t2)
                + "</td>"
                + "<td style=\""
                + td
                + "text-align:right;\">"
                + Fmt.n0(adj)
                + "</td>"
                + "<td style=\""
                + td
                + "text-align:right;\">"
                + Fmt.n0(tg)
                + "</td>"
                + "</tr>\n";
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
