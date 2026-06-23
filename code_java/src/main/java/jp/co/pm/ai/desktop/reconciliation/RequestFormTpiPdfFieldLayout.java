package jp.co.pm.ai.desktop.reconciliation;

import java.text.Normalizer;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** TPI（東レペフ加工品）QR-06-011 PDF 依頼書のフィールド抽出・正規化の正本。 */
public final class RequestFormTpiPdfFieldLayout {

    public static final String META_SOURCE_KIND = "原本種別";
    public static final String META_SOURCE_KIND_TPI_PDF = "TPI_PDF";
    public static final String META_TPI_LAYOUT = "TPI様式";
    public static final String META_SHEET_NAME = "_pdf";

    public static final String LAYOUT_ECOWD = "ECOWD";
    public static final String LAYOUT_PN = "PN";

    private static final Pattern ECOWD_FILE_NAME =
            Pattern.compile("ECOWD.*?[（(]([JＪRＲ\\d０-９\\-－]+[^）)]*)[）)]", Pattern.CASE_INSENSITIVE);
    private static final Pattern PN_FILE_NAME =
            Pattern.compile("後加工.*?[（(](PN\\d{2}-\\d{2})[）)]", Pattern.CASE_INSENSITIVE);

    private static final Pattern JR_BODY =
            Pattern.compile("[ＪJ][ＲR]([\\d０-９]{6}(?:-[\\d０-９]+)?)");
    private static final Pattern PN_BODY = Pattern.compile("(PN\\d{2}-\\d{2})\\s+202[\\d０-９]");
    private static final Pattern DELIVERY_DATE =
            Pattern.compile("年\\s*([\\d０-９]{1,2})\\s*月\\s*([\\d０-９]{1,2})\\s*日\\s*湖南");
    private static final Pattern DOCUMENT_YEAR = Pattern.compile("20([\\d０-９]{2})");
    private static final Pattern CONTRACT_NO = Pattern.compile("X[\\d０-９]{9}");
    private static final Pattern P_NUMBER = Pattern.compile("P[\\d０-９]{9}");
    private static final Pattern USER_LINE =
            Pattern.compile("長岡産業[（(]株[）)]\\s*湖南工場");
    private static final Pattern SC_FEED =
            Pattern.compile("SC[：:]\\s*(\\S+)\\s+投入先[：:]\\s*(\\S+)");
    private static final Pattern FEL_PRODUCT =
            Pattern.compile("(FEL[\\dA-Z\\-]+(?:-EC)?)");

    private RequestFormTpiPdfFieldLayout() {}

    public enum LayoutKind {
        ECOWD,
        PN
    }

    static LayoutKind detectLayout(String fileName, String text) {
        String name = fileName != null ? fileName : "";
        if (name.contains("ECOWD") || name.contains("JR")) {
            return LayoutKind.ECOWD;
        }
        if (name.contains("後加工") || name.contains("PN")) {
            return LayoutKind.PN;
        }
        String body = normalizeText(text);
        if (body.contains("JR屋根") || JR_BODY.matcher(body).find()) {
            return LayoutKind.ECOWD;
        }
        if (PN_BODY.matcher(body).find()) {
            return LayoutKind.PN;
        }
        return LayoutKind.PN;
    }

    static String normalizeText(String text) {
        if (text == null || text.isBlank()) {
            return "";
        }
        return Normalizer.normalize(text, Normalizer.Form.NFKC);
    }

    static String normalizeIraiNo(String raw) {
        if (raw == null || raw.isBlank()) {
            return "";
        }
        String n = toAsciiDigits(Normalizer.normalize(raw, Normalizer.Form.NFKC).strip());
        n = n.replace('Ｊ', 'J').replace('Ｒ', 'R');
        n = n.replaceAll("[－—–]", "-");
        if (n.startsWith("JR")) {
            return n.toUpperCase();
        }
        if (n.matches("R\\d{6}.*")) {
            return ("J" + n).toUpperCase();
        }
        return n.toUpperCase();
    }

    static String parseIraiNoFromFileName(String fileName) {
        if (fileName == null || fileName.isBlank()) {
            return "";
        }
        Matcher ecowd = ECOWD_FILE_NAME.matcher(fileName);
        if (ecowd.find()) {
            String token = ecowd.group(1).strip();
            if (token.toUpperCase().startsWith("JR") || token.startsWith("ＪＲ")) {
                return normalizeIraiNo(token);
            }
            return normalizeIraiNo("JR" + token.replaceFirst("^[JＪRＲ]+", ""));
        }
        Matcher pn = PN_FILE_NAME.matcher(fileName);
        if (pn.find()) {
            return normalizeIraiNo(pn.group(1));
        }
        return "";
    }

    static String parseIraiNoFromText(String text) {
        String body = normalizeText(text);
        Matcher jr = JR_BODY.matcher(body);
        if (jr.find()) {
            return normalizeIraiNo("JR" + jr.group(1));
        }
        Matcher pn = PN_BODY.matcher(body);
        if (pn.find()) {
            return normalizeIraiNo(pn.group(1));
        }
        return "";
    }

    static String resolveIraiNo(String fileName, String text) {
        String fromBody = parseIraiNoFromText(text);
        if (!fromBody.isBlank()) {
            return fromBody;
        }
        return parseIraiNoFromFileName(fileName);
    }

    static int resolveDocumentYear(String text) {
        String body = normalizeText(text);
        Matcher m = DOCUMENT_YEAR.matcher(body);
        int last = 2026;
        while (m.find()) {
            last = Integer.parseInt("20" + toAsciiDigits(m.group(1)));
        }
        return last;
    }

    static String parseDeliveryDate(String text) {
        String body = normalizeText(text);
        Matcher m = DELIVERY_DATE.matcher(body);
        String last = "";
        while (m.find()) {
            int year = resolveDocumentYear(text);
            int month = Integer.parseInt(toAsciiDigits(m.group(1)));
            int day = Integer.parseInt(toAsciiDigits(m.group(2)));
            last = String.format("%04d-%02d-%02d", year, month, day);
        }
        return last;
    }

    static String parseContractNo(String text) {
        String body = normalizeText(text);
        Matcher m = CONTRACT_NO.matcher(body);
        String last = "";
        while (m.find()) {
            last = toAsciiDigits(m.group());
        }
        return last;
    }

    static String parsePNumber(String text) {
        String body = normalizeText(text);
        Matcher m = P_NUMBER.matcher(body);
        String last = "";
        while (m.find()) {
            last = toAsciiDigits(m.group());
        }
        return last;
    }

    static String parseUser(String text) {
        if (USER_LINE.matcher(normalizeText(text)).find()) {
            return "長岡産業（株）湖南工場";
        }
        return "長岡産業（株）湖南工場";
    }

    static String parseFelProductCode(String text) {
        String body = normalizeText(text);
        Matcher m = FEL_PRODUCT.matcher(body);
        String preferred = "";
        while (m.find()) {
            String hit = m.group(1);
            if (hit.endsWith("-EC")) {
                return hit;
            }
            preferred = hit;
        }
        return preferred;
    }

    static String parseScCode(String text) {
        Matcher m = SC_FEED.matcher(normalizeText(text));
        return m.find() ? m.group(1).strip() : "";
    }

    static String parseFeedHint(String text) {
        Matcher m = SC_FEED.matcher(normalizeText(text));
        return m.find() ? m.group(2).strip() : "";
    }

    static String parseEcSide(String text) {
        String body = normalizeText(text);
        if (body.contains("EC（片面）") || body.contains("EC(片面)")) {
            return "EC（片面）";
        }
        if (body.contains("EC（両面）") || body.contains("EC(両面)")) {
            return "EC（両面）";
        }
        return "";
    }

    static String parseProcessingContentEcowd(String text) {
        String body = normalizeText(text);
        StringBuilder sb = new StringBuilder();
        if (body.contains("接続・分割") || body.contains("接続･分割")) {
            appendPart(sb, "接続・分割");
        }
        if (body.contains("JR屋根")) {
            if (body.contains("穴あけ")) {
                appendPart(sb, "JR屋根：EC（片面）穴あけ");
            } else if (body.contains("スリット")) {
                appendPart(sb, "JR屋根：スリット");
            } else {
                appendPart(sb, "JR屋根");
            }
        } else if (body.contains("EC（片面）") && body.contains("穴あけ")) {
            appendPart(sb, "EC（片面）穴あけ");
        }
        if (body.contains("カット品の長さ")) {
            appendPart(sb, "カット品の長さ");
        }
        if (body.contains("ロール品 or カット品") || body.contains("ロール品 or カット品")) {
            appendPart(sb, "ロール品 or カット品");
        }
        return sb.toString();
    }

    static String parseProcessingContentPn(String text) {
        String body = normalizeText(text);
        StringBuilder sb = new StringBuilder();
        String ec = parseEcSide(text);
        if (!ec.isBlank()) {
            appendPart(sb, ec);
        }
        if (body.contains("ロール品 or カット品")) {
            appendPart(sb, "ロール品 or カット品");
        }
        if (body.contains("カット品の長さ")) {
            appendPart(sb, "カット品の長さ");
        }
        return sb.toString();
    }

    static String parseTokkiFromText(String text, String fileName) {
        StringBuilder sb = new StringBuilder();
        String body = normalizeText(text);
        if (fileName != null && fileName.contains("熱融着")) {
            appendTokkiPart(sb, "赤テープ：つなぎありのため熱融着必要");
        }
        if (body.contains("入庫お願い")) {
            String p = parsePNumber(text);
            appendTokkiPart(
                    sb, p.isBlank() ? "入庫お願いします" : "入庫お願いします。『" + p + "』");
        }
        String elRawRemark = parseElRawInputRemark(text);
        if (!elRawRemark.isBlank()) {
            appendTokkiPart(sb, elRawRemark);
        }
        if (body.contains("メールボックス")) {
            appendTokkiPart(
                    sb, "製品ラベルはメールボックスへ。ペフエコード同様2枚貼付（原反側・梱包側）");
        }
        return sb.toString();
    }

    /** PDF ■備考■ 付近の「EL原反は…投入します。」 */
    static String parseElRawInputRemark(String text) {
        if (text == null || text.isBlank()) {
            return "";
        }
        String body = normalizeText(text);
        java.util.regex.Pattern pattern =
                java.util.regex.Pattern.compile("(?:EL|ＥＬ)原反は[^。\\n]+。");
        java.util.regex.Matcher matcher = pattern.matcher(body);
        String last = "";
        while (matcher.find()) {
            last = matcher.group().strip();
        }
        if (last.isBlank()) {
            return "";
        }
        return last.replaceFirst("^[＊\\*]\\s*", "");
    }

    static String translateTpiYoto(LayoutKind kind, String text) {
        String body = normalizeText(text);
        if (kind == LayoutKind.ECOWD || body.contains("JR屋根")) {
            return "JR（屋根）";
        }
        return "V（TPI）";
    }

    static String normalizeTpiLightGrayColor() {
        return "ﾗｲﾄｸﾞﾚｰ";
    }

    static boolean containsLightGrayColor(String text) {
        if (text == null || text.isBlank()) {
            return false;
        }
        String body = normalizeText(text);
        return body.contains("ライトグレー")
                || body.contains("ﾗｲﾄｸﾞﾚｰ")
                || (body.contains("ﾗｲﾄ") && body.contains("ｸﾞﾚ"))
                || (body.contains("ライト") && body.contains("グレ"));
    }

    static String extractTpiLightGrayColorIfPresent(String text) {
        return containsLightGrayColor(text) ? normalizeTpiLightGrayColor() : "";
    }

    static String toAsciiDigits(String s) {
        if (s == null) {
            return "";
        }
        StringBuilder out = new StringBuilder(s.length());
        for (char c : s.toCharArray()) {
            if (c >= '０' && c <= '９') {
                out.append((char) ('0' + (c - '０')));
            } else {
                out.append(c);
            }
        }
        return out.toString();
    }

    private static void appendPart(StringBuilder sb, String part) {
        if (part == null || part.isBlank()) {
            return;
        }
        if (!sb.isEmpty()) {
            sb.append(", ");
        }
        sb.append(part.strip());
    }

    private static void appendTokkiPart(StringBuilder sb, String part) {
        if (part == null || part.isBlank()) {
            return;
        }
        if (!sb.isEmpty()) {
            sb.append('　');
        }
        sb.append(part.strip());
    }
}
