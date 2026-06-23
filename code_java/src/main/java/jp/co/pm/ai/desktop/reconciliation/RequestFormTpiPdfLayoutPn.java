package jp.co.pm.ai.desktop.reconciliation;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import jp.co.pm.ai.desktop.reconciliation.RequestFormTpiPdfFieldLayout.LayoutKind;

/** 後加工 / PN 系 TPI PDF から rawMap を組み立てる。 */
final class RequestFormTpiPdfLayoutPn {

    private static final Pattern PRODUCT_HINMEI =
            Pattern.compile("加工製品\\s+(?:[①１1]\\s+)?(\\S+)");
    private static final Pattern RAW_LINE =
            Pattern.compile("投入原反\\s+(?:[①１1]\\s+)?(\\S+)\\s+（([^）)]+)）\\s+([\\d,]+)\\s+(\\d{1,2}/\\d{1,2})");
    private static final Pattern QTY_METERS =
            Pattern.compile("([\\d,]+)\\s*[ｍm]\\s*\\n?\\s*\\d+\\s+([\\d,]+)");
    private static final Pattern QTY_SIMPLE = Pattern.compile("加工賃\\s*([\\d,]+)\\s*[ｍm]");

    private RequestFormTpiPdfLayoutPn() {}

    static Map<String, String> buildRawMap(String fileName, String text) {
        Map<String, String> raw = new LinkedHashMap<>();
        String body = RequestFormTpiPdfFieldLayout.normalizeText(text);

        raw.put("依頼Ｎｏ", RequestFormTpiPdfFieldLayout.resolveIraiNo(fileName, text));
        raw.put("希望納期", RequestFormTpiPdfFieldLayout.parseDeliveryDate(text));
        raw.put("ユーザー", RequestFormTpiPdfFieldLayout.parseUser(text));
        raw.put("契約Ｎｏ", RequestFormOriginalExtractor.resolveContractNoFromOriginalCell(
                RequestFormTpiPdfFieldLayout.parseContractNo(text)));

        Matcher product = PRODUCT_HINMEI.matcher(body);
        if (product.find()) {
            raw.put("品名", product.group(1));
        }

        String fel = RequestFormTpiPdfFieldLayout.parseFelProductCode(text);
        if (!fel.isBlank()) {
            raw.put("製品", fel);
        }

        Matcher rawLine = RAW_LINE.matcher(body);
        if (rawLine.find()) {
            raw.put("原反品名", rawLine.group(1));
            raw.put("原反", rawLine.group(1) + "（" + rawLine.group(2).strip() + "）");
            raw.put("原反数量", RequestFormTpiPdfFieldLayout.toAsciiDigits(rawLine.group(3)).replace(",", ""));
            raw.put("投入日", rawLine.group(4));
        }

        String qty = parseQuantity(body);
        if (!qty.isBlank()) {
            raw.put("数量1", qty);
        }

        String feed = RequestFormTpiPdfFieldLayout.parseFeedHint(text);
        if (!feed.isBlank()) {
            raw.put("投入場所", feed);
        }

        String processing = RequestFormTpiPdfFieldLayout.parseProcessingContentPn(text);
        raw.put("加工内容", processing);
        raw.put("ＥＣ面", RequestFormTpiPdfFieldLayout.parseEcSide(text));
        raw.put("用途", RequestFormTpiPdfFieldLayout.translateTpiYoto(LayoutKind.PN, text));

        String tokki2 = RequestFormTpiPdfFieldLayout.parseTokkiFromText(text, fileName);
        if (!tokki2.isBlank()) {
            raw.put("特記事項2", tokki2);
        }

        fillMeta(raw, fileName);
        return raw;
    }

    private static String parseQuantity(String body) {
        Matcher simple = QTY_SIMPLE.matcher(body);
        if (simple.find()) {
            return RequestFormTpiPdfFieldLayout.toAsciiDigits(simple.group(1)).replace(",", "");
        }
        Matcher m = Pattern.compile("([\\d,]+)\\s*[ｍm]\\s+\\d+\\s+[\\d,]+").matcher(body);
        if (m.find()) {
            return RequestFormTpiPdfFieldLayout.toAsciiDigits(m.group(1)).replace(",", "");
        }
        Matcher inline = Pattern.compile("EC（片面）\\s+\\d+\\s+ロール品[^\\d]*([\\d,]+)").matcher(body);
        if (inline.find()) {
            return RequestFormTpiPdfFieldLayout.toAsciiDigits(inline.group(1)).replace(",", "");
        }
        return "";
    }

    private static void fillMeta(Map<String, String> raw, String fileName) {
        raw.put("原本ファイル名", fileName != null ? fileName : "");
        raw.put("原本シート名", RequestFormTpiPdfFieldLayout.META_SHEET_NAME);
        raw.put(RequestFormTpiPdfFieldLayout.META_SOURCE_KIND, RequestFormTpiPdfFieldLayout.META_SOURCE_KIND_TPI_PDF);
        raw.put(RequestFormTpiPdfFieldLayout.META_TPI_LAYOUT, RequestFormTpiPdfFieldLayout.LAYOUT_PN);
    }
}
