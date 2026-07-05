package jp.co.pm.ai.desktop.reconciliation;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import jp.co.pm.ai.desktop.reconciliation.RequestFormTpiPdfFieldLayout.LayoutKind;

/** ECOWD / JR 系 TPI PDF から rawMap を組み立てる。 */
final class RequestFormTpiPdfLayoutEcowd {

    private static final String DEFAULT_GRADE_WHEN_BLANK = "F-A";

    private static final Pattern PRODUCT_HEADER =
            Pattern.compile("加工製品\\s+(?:[①１1]\\s+)?(\\d+)\\s+(\\S+)\\s+(\\S+)");
    /** 製品行: 「95 190m」（幅+長さ）または「95 950m」（長さ+数量）。 */
    private static final Pattern DIM_LINE_WITH_M =
            Pattern.compile("(?:^|\\n)\\s*(\\d{1,4})\\s+([\\d,０-９]+)\\s*m\\b", Pattern.MULTILINE);
    /** 製品行: 長さ＋色＋数量が分離するケース（JR260603 系）。 */
    private static final Pattern LENGTH_COLOR_QTY =
            Pattern.compile(
                    "(\\d{1,4})\\s+(?:ﾗｲﾄ|ライト)[\\s\\S]{0,24}?(?:ｸﾞﾚ|グレ)[\\s\\S]{0,12}?([\\d,０-９]+)\\s*m\\b");
    private static final Pattern PRODUCT_QTY =
            Pattern.compile("(?:^|\\n)\\s*([\\d,０-９]+)\\s*m\\b", Pattern.MULTILINE | Pattern.CASE_INSENSITIVE);
    private static final Pattern RAW_HEADER =
            Pattern.compile("投入原反\\s+(?:[①１1②２2]\\s+)?(\\S+)\\s+(.+)");
    /** 投入原反の数量行: 「100 ﾗｲﾄｸﾞﾚｰ 500m 6/10」 */
    private static final Pattern RAW_COLOR_QTY_DATE =
            Pattern.compile(
                    "(\\d{1,4})\\s+(?:ﾗｲﾄ|ライト)[\\s\\S]{0,24}?(?:ｸﾞﾚ|グレ)[\\s\\S]{0,12}?([\\d,０-９]+)\\s*m\\s+(\\d{1,2}/\\d{1,2})");
    /** 投入原反の数量行（色なし）: 「100 1,000m 6/10」 */
    private static final Pattern RAW_QTY_DATE =
            Pattern.compile(
                    "(\\d{1,4})\\s+([\\d,０-９]+)\\s*m\\s+(\\d{1,2}/\\d{1,2})");
    private static final Pattern KAKOCHIN =
            Pattern.compile("加工賃\\s*([\\d,]+)\\s*[ｍm]");
    private static final Pattern NOTE_LINE =
            Pattern.compile("^[＊\\*](.+)$", Pattern.MULTILINE);

    private RequestFormTpiPdfLayoutEcowd() {}

    static Map<String, String> buildRawMap(String fileName, String text) {
        Map<String, String> raw = new LinkedHashMap<>();
        String body = RequestFormTpiPdfFieldLayout.normalizeText(text);

        raw.put("依頼Ｎｏ", RequestFormTpiPdfFieldLayout.resolveIraiNo(fileName, text));
        raw.put("希望納期", RequestFormTpiPdfFieldLayout.parseDeliveryDate(text));
        raw.put("ユーザー", RequestFormTpiPdfFieldLayout.parseUser(text));
        raw.put("契約Ｎｏ", RequestFormOriginalExtractor.resolveContractNoFromOriginalCell(
                RequestFormTpiPdfFieldLayout.parseContractNo(text)));

        fillProduct(raw, body);
        fillRawMaterial(raw, body);
        fillMeta(raw, fileName, text);

        String processing = RequestFormTpiPdfFieldLayout.parseProcessingContentEcowd(text);
        raw.put("加工内容", processing);
        raw.put("ＥＣ面", RequestFormTpiPdfFieldLayout.parseEcSide(text));
        raw.put("用途", RequestFormTpiPdfFieldLayout.translateTpiYoto(LayoutKind.ECOWD, text));

        Matcher kako = KAKOCHIN.matcher(body);
        if (kako.find()) {
            String meters = RequestFormTpiPdfFieldLayout.toAsciiDigits(kako.group(1)).replace(",", "");
            raw.put("加工賃", meters + " ｍ");
            if (!raw.containsKey("数量1") || raw.get("数量1").isBlank()) {
                raw.put("数量1", meters);
            }
        }

        String tokki1 = extractHeaderNote(body, fileName);
        if (!tokki1.isBlank()) {
            raw.put("特記事項1", tokki1);
        }
        String tokki2 = RequestFormTpiPdfFieldLayout.parseTokkiFromText(text, fileName);
        if (!tokki2.isBlank()) {
            raw.put("特記事項2", tokki2);
        }

        return raw;
    }

    private static boolean looksLikeEcowdPartNumberType(String typeToken) {
        return typeToken != null && typeToken.matches("A\\d{2}W");
    }

    private static Matcher findProductHeader(String body) {
        Matcher header = PRODUCT_HEADER.matcher(body);
        Matcher fallback = null;
        while (header.find()) {
            String hinmei = header.group(1);
            String part = header.group(2);
            if (hinmei.length() >= 4 || part.startsWith("R") || part.startsWith("FEL")) {
                return header;
            }
            if (fallback == null) {
                fallback = header;
            }
        }
        return fallback;
    }

    private static void fillProduct(Map<String, String> raw, String body) {
        Matcher header = findProductHeader(body);
        if (header == null) {
            return;
        }
        String hinmeiToken = header.group(1);
        String part = header.group(2);
        // A05W 系は先頭数字が品番（30020）。R10W 系は先頭数字が品名コード（40040）。
        if (!hinmeiToken.matches("\\d+") || !looksLikeEcowdPartNumberType(part)) {
            raw.put("品名", hinmeiToken);
        }
        String headerThird = header.group(3);

        ProductDims dims = parseProductDims(body, header.start(), header.end(), headerThird);

        raw.put("製品", JuchuSheetColumnLayout.buildSpecName(part, dims.type, dims.width, dims.length));
        String color = dims.color;
        if (color.isBlank()) {
            color =
                    RequestFormTpiPdfFieldLayout.extractTpiLightGrayColorIfPresent(
                            windowAroundHeader(body, header.start(), header.end()));
        }
        raw.put("色1", RequestFormOriginalExtractor.resolveOriginalColor(color));
        raw.put("梱-等1", dims.grade.isBlank() ? DEFAULT_GRADE_WHEN_BLANK : dims.grade);
        if (!dims.quantity.isBlank()) {
            raw.put("数量1", dims.quantity);
        }
    }

    private static String windowAroundHeader(String body, int headerStart, int headerEnd) {
        int begin = Math.max(0, headerStart - 400);
        int end = Math.min(body.length(), headerEnd + 400);
        return body.substring(begin, end);
    }

    private static ProductDims parseProductDims(
            String body, int headerStart, int headerEnd, String headerThird) {
        ProductDims dims = new ProductDims();
        dims.type = headerThird;

        String forward = body.substring(headerEnd, Math.min(body.length(), headerEnd + 400));
        String backward =
                body.substring(Math.max(0, headerStart - 400), headerStart);
        String around = windowAroundHeader(body, headerStart, headerEnd);

        LengthColorQtyMatch lengthColorForward = findLengthColorQty(forward, false);
        LengthColorQtyMatch lengthColorBackward = findLengthColorQty(backward, true);
        LengthColorQtyMatch lengthColor =
                lengthColorForward != null ? lengthColorForward : lengthColorBackward;
        if (lengthColor != null) {
            dims.length = lengthColor.length();
            dims.width = headerThird;
            dims.quantity =
                    RequestFormTpiPdfFieldLayout.parseQuantityDigits(lengthColor.quantity());
            dims.color =
                    RequestFormTpiPdfFieldLayout.extractTpiLightGrayColorIfPresent(
                            lengthColor.snippet());
            return dims;
        }

        DimLineMatch dimForward = findDimLine(forward, false);
        DimLineMatch dimBackward = findDimLine(backward, true);
        DimLineMatch dimLine = dimForward != null ? dimForward : dimBackward;
        if (dimLine == null) {
            dimLine = findDimLine(around, true);
        }
        if (dimLine != null) {
            applyDimLine(dims, headerThird, dimLine);
            if (dims.color.isBlank()) {
                dims.color =
                        RequestFormTpiPdfFieldLayout.extractTpiLightGrayColorIfPresent(
                                dimLine.snippet());
            }
            return dims;
        }

        dims.width = headerThird;
        dims.length = "";
        dims.quantity = findProductQuantity(around, dims.length);
        dims.color = RequestFormTpiPdfFieldLayout.extractTpiLightGrayColorIfPresent(around);
        return dims;
    }

    private static LengthColorQtyMatch findLengthColorQty(String scope, boolean preferLastMatch) {
        Matcher matcher = LENGTH_COLOR_QTY.matcher(scope);
        LengthColorQtyMatch chosen = null;
        int chosenStart = preferLastMatch ? -1 : Integer.MAX_VALUE;
        while (matcher.find()) {
            int start = matcher.start();
            if (preferLastMatch ? start > chosenStart : start < chosenStart) {
                chosenStart = start;
                chosen =
                        new LengthColorQtyMatch(
                                matcher.group(1), matcher.group(2), matcher.group(0));
            }
        }
        return chosen;
    }

    private static DimLineMatch findDimLine(String scope, boolean preferLastMatch) {
        Matcher matcher = DIM_LINE_WITH_M.matcher(scope);
        DimLineMatch chosen = null;
        int chosenStart = preferLastMatch ? -1 : Integer.MAX_VALUE;
        while (matcher.find()) {
            int start = matcher.start();
            if (preferLastMatch ? start > chosenStart : start < chosenStart) {
                chosenStart = start;
                chosen =
                        new DimLineMatch(
                                matcher.group(1),
                                matcher.group(2),
                                matcher.group(0));
            }
        }
        return chosen;
    }

    /** ECOWD 製品2行目 {@code 長さ 数量m} は常に長さ・数量。幅はヘッダ3列目（例: 870）。 */
    private static void applyDimLine(ProductDims dims, String headerThird, DimLineMatch dimLine) {
        dims.width = headerThird;
        dims.length = dimLine.first();
        dims.quantity = RequestFormTpiPdfFieldLayout.parseQuantityDigits(dimLine.secondMeters());
        dims.type = headerThird;
    }

    private static String findProductQuantity(String productBlock, String lengthWithUnit) {
        String lengthDigits =
                lengthWithUnit == null || lengthWithUnit.isBlank()
                        ? ""
                        : RequestFormTpiPdfFieldLayout.toAsciiDigits(lengthWithUnit)
                                .replace("m", "")
                                .replace("ｍ", "");
        Matcher qty = PRODUCT_QTY.matcher(productBlock);
        while (qty.find()) {
            String candidate =
                    RequestFormTpiPdfFieldLayout.parseQuantityDigits(qty.group(1));
            if (!candidate.isBlank() && !candidate.equals(lengthDigits)) {
                return candidate;
            }
        }
        return "";
    }

    private static void fillRawMaterial(Map<String, String> raw, String body) {
        Matcher header = RAW_HEADER.matcher(body);
        if (!header.find()) {
            return;
        }
        int firstHeaderStart = header.start();
        int windowEnd = header.end() + 800;

        String hinmei = header.group(1);
        String tail = header.group(2).strip();
        raw.put("原反品名", hinmei);
        String[] tokens = tail.split("\\s+");
        String part = tokens.length > 0 ? tokens[0] : "";
        String type = tokens.length > 1 ? tokens[1] : "";
        String width = tokens.length > 2 ? tokens[2] : "";
        String length = tokens.length > 3 ? tokens[3] : "";
        raw.put("原反", JuchuSheetColumnLayout.buildSpecName(part, type, width, length));

        while (header.find()) {
            windowEnd = Math.max(windowEnd, header.end() + 800);
        }

        int searchFrom = Math.max(0, firstHeaderStart - 400);
        int searchTo = Math.min(body.length(), windowEnd);
        List<RawMaterialLine> lines = findAllRawMaterialLines(body.substring(searchFrom, searchTo));

        List<String> qtys = new ArrayList<>();
        List<String> dates = new ArrayList<>();
        for (RawMaterialLine line : lines) {
            if (!line.quantity.isBlank()
                    && (qtys.isEmpty() || !qtys.get(qtys.size() - 1).equals(line.quantity))) {
                qtys.add(line.quantity);
            }
            if (!line.inputDate.isBlank()
                    && (dates.isEmpty() || !dates.get(dates.size() - 1).equals(line.inputDate))) {
                dates.add(line.inputDate);
            }
            if (!line.color.isBlank() && !raw.containsKey("原反色")) {
                raw.put("原反色", line.color);
            }
        }
        if (!qtys.isEmpty()) {
            raw.put("原反数量", String.join("\n", qtys));
        }
        if (!dates.isEmpty()) {
            raw.put("投入日", String.join("\n", dates));
        }
        if (!raw.containsKey("原反色")) {
            raw.put(
                    "原反色",
                    RequestFormOriginalExtractor.resolveOriginalColor(
                            RequestFormTpiPdfFieldLayout.extractTpiLightGrayColorIfPresent(body)));
        }
    }

    /** 投入原反ブロック内の数量・投入日行を出現順に収集（同一位置の重複は除外）。 */
    private static List<RawMaterialLine> findAllRawMaterialLines(String rawWindow) {
        TreeMap<Integer, RawMaterialLine> byStart = new TreeMap<>();
        Matcher detail = RAW_COLOR_QTY_DATE.matcher(rawWindow);
        while (detail.find()) {
            byStart.put(
                    detail.start(),
                    new RawMaterialLine(
                            RequestFormTpiPdfFieldLayout.parseQuantityDigits(detail.group(2)),
                            detail.group(3),
                            RequestFormOriginalExtractor.resolveOriginalColor(
                                    RequestFormTpiPdfFieldLayout.extractTpiLightGrayColorIfPresent(
                                            detail.group(0)))));
        }
        Matcher qtyDate = RAW_QTY_DATE.matcher(rawWindow);
        while (qtyDate.find()) {
            int start = qtyDate.start();
            if (overlapsExistingRawLine(byStart, start)) {
                continue;
            }
            byStart.put(
                    start,
                    new RawMaterialLine(
                            RequestFormTpiPdfFieldLayout.parseQuantityDigits(qtyDate.group(2)),
                            qtyDate.group(3),
                            ""));
        }
        return new ArrayList<>(byStart.values());
    }

    private static boolean overlapsExistingRawLine(TreeMap<Integer, RawMaterialLine> byStart, int start) {
        Integer floor = byStart.floorKey(start);
        if (floor != null && start - floor < 24) {
            return true;
        }
        Integer ceiling = byStart.ceilingKey(start);
        return ceiling != null && ceiling - start < 24;
    }

    private record RawMaterialLine(String quantity, String inputDate, String color) {}

    private static void fillMeta(Map<String, String> raw, String fileName, String text) {
        raw.put("原本ファイル名", fileName != null ? fileName : "");
        raw.put("原本シート名", RequestFormTpiPdfFieldLayout.META_SHEET_NAME);
        raw.put(RequestFormTpiPdfFieldLayout.META_SOURCE_KIND, RequestFormTpiPdfFieldLayout.META_SOURCE_KIND_TPI_PDF);
        raw.put(RequestFormTpiPdfFieldLayout.META_TPI_LAYOUT, RequestFormTpiPdfFieldLayout.LAYOUT_ECOWD);
    }

    private static String extractHeaderNote(String body, String fileName) {
        if (fileName != null && fileName.contains("熱融着")) {
            return "赤テープ：つなぎありのため熱融着必要";
        }
        for (String line : body.split("\\R")) {
            String t = line.strip();
            if (t.contains("赤テープ") || t.contains("熱融着")) {
                return t;
            }
        }
        Matcher note = NOTE_LINE.matcher(body);
        if (note.find() && note.group(1).contains("穴あけ")) {
            return "＊" + note.group(1).strip();
        }
        return "";
    }

    private static final class ProductDims {
        String type = "";
        String width = "";
        String length = "";
        String quantity = "";
        String color = "";
        String grade = "";
    }

    private record LengthColorQtyMatch(String length, String quantity, String snippet) {}

    private record DimLineMatch(String first, String secondMeters, String snippet) {}
}
