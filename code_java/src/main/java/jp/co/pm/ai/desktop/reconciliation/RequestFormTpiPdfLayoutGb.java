package jp.co.pm.ai.desktop.reconciliation;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** 古河原反スライス加工依頼 PDF（例: GB.pdf）から rawMap を組み立てる。 */
final class RequestFormTpiPdfLayoutGb {

    private static final Pattern GB_ORDER_HEADER =
            Pattern.compile(
                    "(?:依頼No[,.]?|No\\.)(G\\s*B\\s*[\\d０-９]{4,6})",
                    Pattern.CASE_INSENSITIVE);
    private static final Pattern GB_DELIVERY =
            Pattern.compile("納期\\s*([\\d０-９]{1,2})\\s*月\\s*([\\d０-９]{1,2})\\s*日");
    private static final Pattern GB_DOCUMENT_DATE =
            Pattern.compile("20([\\d０-９]{2})\\s*年\\s*([\\d０-９]{1,2})\\s*月\\s*([\\d０-９]{1,2})\\s*日");
    private static final Pattern GB_HB_PRODUCT_ROW =
            Pattern.compile(
                    "(HB\\d+GB)\\D*?([\\d.]+)mm\\s+(\\d{3,4})\\s+(\\d+)\\s+(\\d+)\\s+(\\d+)",
                    Pattern.CASE_INSENSITIVE);
    /** 例: {@code 15025 NR28 BK0 1560 50 1 50} / {@code 15025 AW07 400 50 1 50} */
    private static final Pattern GB_GENERIC_PRODUCT_ROW =
            Pattern.compile(
                    "(\\d{4,5})\\s+([A-Z0-9]+(?:\\s+[A-Z0-9]+)?)\\s+(\\d{3,4})\\s+(\\d+)\\s+(\\d+)\\s+(\\d+)",
                    Pattern.CASE_INSENSITIVE);
    /** {@code 古河原反} の {@code 原反} 部分と区別する（{@code (?<![河])}）。 */
    private static final Pattern GB_RAW_SECTION_START =
            Pattern.compile("(?<![河])原\\s*反(?:\\s*投入)?|投入原反");

    private enum GbTableSection {
        PRODUCT,
        RAW
    }

    private RequestFormTpiPdfLayoutGb() {}

    /** 1 PDF に複数依頼がある場合は依頼ごとに 1 rawMap。 */
    static List<Map<String, String>> buildAllRawMaps(String fileName, String text) {
        List<String> blocks = splitOrderBlocks(text);
        if (blocks.isEmpty()) {
            return List.of(buildRawMap(fileName, text));
        }
        List<Map<String, String>> out = new ArrayList<>(blocks.size());
        for (String block : blocks) {
            out.add(buildRawMap(fileName, block));
        }
        return out;
    }

    static Map<String, String> buildRawMap(String fileName, String text) {
        Map<String, String> raw = new LinkedHashMap<>();
        raw.put("依頼Ｎｏ", RequestFormTpiPdfFieldLayout.resolveIraiNo(fileName, text));
        raw.put("希望納期", parseGbDeliveryDate(text));
        raw.put("ユーザー", parseGbUser(text));
        raw.put("用途", "小口加工");
        // GB 依頼 PDF には受注フォーム相当の加工内容欄が無く、タイトルの「スライス」「巻き返し」等は転記しない。

        List<String> hinmeis = new ArrayList<>();
        List<String> specs = new ArrayList<>();
        List<String> qtys = new ArrayList<>();
        List<String> colors = new ArrayList<>();

        String normalized = RequestFormTpiPdfFieldLayout.normalizeText(text);
        HbGbTables hbTables = parseHbTables(normalized);
        appendHbRows(hbTables.products(), hinmeis, specs, qtys);
        appendHbRawMaterialRows(hbTables.rawMaterials(), raw);
        GenericGbTables genericTables = parseGenericTables(normalized);
        appendGenericRows(genericTables.products(), hinmeis, specs, qtys, colors);
        appendGenericRawMaterialRows(genericTables.rawMaterials(), raw);

        if (!hinmeis.isEmpty()) {
            raw.put("品名", String.join("\n", hinmeis));
            raw.put("製品", String.join("\n", specs));
            raw.put("数量1", String.join("\n", qtys));
            if (colors.stream().anyMatch(color -> color != null && !color.isBlank())) {
                raw.put("色1", String.join("\n", colors));
            }
        }

        fillMeta(raw, fileName);
        return raw;
    }

    private record GenericGbRow(String part, String typeRaw, String width, String length, String rollQty) {}

    private record GenericGbTables(List<GenericGbRow> products, List<GenericGbRow> rawMaterials) {}

    private record HbGbRow(String part, String thicknessMm, String width, String length, String rollQty) {}

    private record HbGbTables(List<HbGbRow> products, List<HbGbRow> rawMaterials) {}

    /**
     * PDF テキスト抽出では HB 行も {@code 原 反} 見出しより先に来ることがある。
     * 見出し直前の HB 行のみを投入原反とみなし、それより前を製品とする。
     */
    static HbGbTables parseHbTables(String text) {
        List<HbGbRow> products = new ArrayList<>();
        List<HbGbRow> rawMaterials = new ArrayList<>();
        if (text == null || text.isBlank()) {
            return new HbGbTables(products, rawMaterials);
        }
        String normalized = RequestFormTpiPdfFieldLayout.normalizeText(text);
        String[] lines = normalized.split("\\R");
        int rawHeaderLine = findRawSectionHeaderLineIndex(lines);
        List<HbLineMatch> matches = new ArrayList<>();
        for (int i = 0; i < lines.length; i++) {
            Matcher matcher = GB_HB_PRODUCT_ROW.matcher(collapseWhitespace(lines[i]));
            if (matcher.find()) {
                matches.add(new HbLineMatch(i, matcher));
            }
        }
        if (matches.isEmpty()) {
            return new HbGbTables(products, rawMaterials);
        }
        int productLabelLine = findProductLabelLineIndex(lines);
        if (productLabelLine >= 0) {
            int inputLabelLine = findInputLabelLineIndex(lines);
            for (HbLineMatch match : matches) {
                int lineIndex = match.lineIndex();
                if (lineIndex < productLabelLine) {
                    products.add(toHbRow(match.matcher()));
                } else if (inputLabelLine < 0 || lineIndex < inputLabelLine) {
                    rawMaterials.add(toHbRow(match.matcher()));
                }
            }
            return new HbGbTables(products, rawMaterials);
        }
        if (rawHeaderLine < 0) {
            assignHbRowsBySectionMarkers(lines, products, rawMaterials);
            return new HbGbTables(products, rawMaterials);
        }
        Integer rawLineIndex = findRawMaterialDataLineIndex(lines, rawHeaderLine, GB_HB_PRODUCT_ROW);
        for (HbLineMatch match : matches) {
            if (rawLineIndex != null && match.lineIndex() == rawLineIndex) {
                rawMaterials.add(toHbRow(match.matcher()));
            } else if (match.lineIndex() < rawHeaderLine) {
                products.add(toHbRow(match.matcher()));
            }
        }
        return new HbGbTables(products, rawMaterials);
    }

    private record HbLineMatch(int lineIndex, Matcher matcher) {}

    private static HbGbRow toHbRow(Matcher row) {
        return new HbGbRow(
                row.group(1),
                RequestFormTpiPdfFieldLayout.toAsciiDigits(row.group(2)),
                RequestFormTpiPdfFieldLayout.toAsciiDigits(row.group(3)),
                RequestFormTpiPdfFieldLayout.toAsciiDigits(row.group(4)),
                RequestFormTpiPdfFieldLayout.toAsciiDigits(row.group(5)));
    }

    /**
     * PDF テキスト抽出では数値行が {@code 原 反} 見出しより先に来ることがある。
     * 見出し直前の数値行のみを投入原反とみなし、それより前を製品とする。
     */
    static GenericGbTables parseGenericTables(String text) {
        List<GenericGbRow> products = new ArrayList<>();
        List<GenericGbRow> rawMaterials = new ArrayList<>();
        if (text == null || text.isBlank()) {
            return new GenericGbTables(products, rawMaterials);
        }
        String normalized = RequestFormTpiPdfFieldLayout.normalizeText(text);
        String[] lines = normalized.split("\\R");
        int rawHeaderLine = findRawSectionHeaderLineIndex(lines);
        List<LineMatch> matches = new ArrayList<>();
        for (int i = 0; i < lines.length; i++) {
            Matcher matcher = GB_GENERIC_PRODUCT_ROW.matcher(collapseWhitespace(lines[i]));
            if (matcher.find()) {
                matches.add(new LineMatch(i, matcher));
            }
        }
        if (matches.isEmpty()) {
            return new GenericGbTables(products, rawMaterials);
        }
        if (rawHeaderLine < 0) {
            for (LineMatch match : matches) {
                products.add(toGenericRow(match.matcher()));
            }
            return new GenericGbTables(products, rawMaterials);
        }
        Integer rawLineIndex = findRawMaterialDataLineIndex(lines, rawHeaderLine, GB_GENERIC_PRODUCT_ROW);
        for (LineMatch match : matches) {
            if (rawLineIndex != null && match.lineIndex() == rawLineIndex) {
                rawMaterials.add(toGenericRow(match.matcher()));
            } else if (match.lineIndex() < rawHeaderLine) {
                products.add(toGenericRow(match.matcher()));
            }
        }
        return new GenericGbTables(products, rawMaterials);
    }

    private record LineMatch(int lineIndex, Matcher matcher) {}

    private static GenericGbRow toGenericRow(Matcher row) {
        return new GenericGbRow(
                row.group(1),
                row.group(2).trim(),
                RequestFormTpiPdfFieldLayout.toAsciiDigits(row.group(3)),
                RequestFormTpiPdfFieldLayout.toAsciiDigits(row.group(4)),
                RequestFormTpiPdfFieldLayout.toAsciiDigits(row.group(5)));
    }

    private static void assignHbRowsBySectionMarkers(
            String[] lines, List<HbGbRow> products, List<HbGbRow> rawMaterials) {
        GbTableSection section = GbTableSection.PRODUCT;
        for (String line : lines) {
            String collapsed = collapseWhitespace(line);
            if (collapsed.isBlank()) {
                continue;
            }
            GbTableSection marker = resolveGbTableSectionMarker(collapsed);
            if (marker != null && !GB_HB_PRODUCT_ROW.matcher(collapsed).find()) {
                section = marker;
                continue;
            }
            Matcher matcher = GB_HB_PRODUCT_ROW.matcher(collapsed);
            if (!matcher.find()) {
                continue;
            }
            HbGbRow row = toHbRow(matcher);
            if (resolveGbRowSectionFromLine(collapsed, section) == GbTableSection.RAW) {
                rawMaterials.add(row);
            } else {
                products.add(row);
            }
        }
    }

    private static GbTableSection resolveGbTableSectionMarker(String collapsed) {
        if (GB_RAW_SECTION_START.matcher(collapsed).find()) {
            return GbTableSection.RAW;
        }
        if (collapsed.matches("投\\s*入(?:原反)?\\s*")) {
            return GbTableSection.RAW;
        }
        if (collapsed.matches("製\\s*品\\s*")) {
            return GbTableSection.PRODUCT;
        }
        return null;
    }

    private static GbTableSection resolveGbRowSectionFromLine(
            String collapsed, GbTableSection currentSection) {
        if (collapsed.matches("投\\s*入(?:原反)?\\s+.*")) {
            return GbTableSection.RAW;
        }
        if (collapsed.matches("製\\s*品\\s+.*")) {
            return GbTableSection.PRODUCT;
        }
        return currentSection;
    }

    static int findProductLabelLineIndex(String[] lines) {
        for (int i = 0; i < lines.length; i++) {
            String collapsed = collapseWhitespace(lines[i]);
            if (collapsed.matches("製\\s*品\\s*")) {
                return i;
            }
        }
        return -1;
    }

    static int findInputLabelLineIndex(String[] lines) {
        int inputLine = -1;
        for (int i = 0; i < lines.length; i++) {
            String collapsed = collapseWhitespace(lines[i]);
            if (collapsed.isBlank()) {
                continue;
            }
            if (collapsed.matches("投\\s*入\\s*")) {
                inputLine = i;
            }
        }
        return inputLine;
    }

    static int findRawSectionHeaderLineIndex(String[] lines) {
        int rawLine = -1;
        int inputLine = -1;
        for (int i = 0; i < lines.length; i++) {
            String collapsed = collapseWhitespace(lines[i]);
            if (collapsed.isBlank()) {
                continue;
            }
            if (GB_RAW_SECTION_START.matcher(collapsed).find()
                    || collapsed.matches("(?<![河])原\\s*反")) {
                rawLine = i;
            }
            if (collapsed.matches("投\\s*入\\s*")) {
                inputLine = i;
            }
        }
        if (rawLine >= 0) {
            return rawLine;
        }
        return inputLine;
    }

    private static Integer findRawMaterialDataLineIndex(String[] lines, int rawHeaderLine, Pattern rowPattern) {
        for (int i = rawHeaderLine - 1; i >= 0; i--) {
            String collapsed = collapseWhitespace(lines[i]);
            if (collapsed.isBlank()) {
                continue;
            }
            if (rowPattern.matcher(collapsed).find()) {
                return i;
            }
        }
        return null;
    }

    static String extractProductSectionBody(String body) {
        if (body == null || body.isBlank()) {
            return "";
        }
        Matcher rawSection = GB_RAW_SECTION_START.matcher(body);
        if (rawSection.find()) {
            return body.substring(0, rawSection.start()).strip();
        }
        return body.strip();
    }

    static String extractRawMaterialSectionBody(String body) {
        if (body == null || body.isBlank()) {
            return "";
        }
        Matcher rawSection = GB_RAW_SECTION_START.matcher(body);
        if (rawSection.find()) {
            return body.substring(rawSection.start()).strip();
        }
        return "";
    }

    static List<String> splitOrderBlocks(String text) {
        String body = text != null ? text : "";
        Matcher header = GB_ORDER_HEADER.matcher(body);
        List<Integer> starts = new ArrayList<>();
        while (header.find()) {
            starts.add(header.start());
        }
        if (starts.isEmpty()) {
            return List.of();
        }
        List<String> blocks = new ArrayList<>(starts.size());
        for (int i = 0; i < starts.size(); i++) {
            int start = starts.get(i);
            int end = i + 1 < starts.size() ? starts.get(i + 1) : body.length();
            blocks.add(body.substring(start, end));
        }
        return blocks;
    }

    static String parseGbDeliveryDate(String text) {
        String body = RequestFormTpiPdfFieldLayout.normalizeText(text);
        int year = RequestFormTpiPdfFieldLayout.resolveDocumentYear(text);
        Matcher doc = GB_DOCUMENT_DATE.matcher(body);
        if (doc.find()) {
            year = Integer.parseInt("20" + RequestFormTpiPdfFieldLayout.toAsciiDigits(doc.group(1)));
        }
        Matcher delivery = GB_DELIVERY.matcher(body);
        if (!delivery.find()) {
            return "";
        }
        int month = Integer.parseInt(RequestFormTpiPdfFieldLayout.toAsciiDigits(delivery.group(1)));
        int day = Integer.parseInt(RequestFormTpiPdfFieldLayout.toAsciiDigits(delivery.group(2)));
        return String.format("%04d-%02d-%02d", year, month, day);
    }

    private static String parseGbUser(String text) {
        String body = RequestFormTpiPdfFieldLayout.normalizeText(text);
        if (body.contains("長岡産業")) {
            return "長岡産業（株）湖南工場";
        }
        return RequestFormTpiPdfFieldLayout.parseUser(text);
    }

    /** 厚み4.5mm → 45、厚み9mm → 90（受注 spec のタイプ桁に合わせる）。 */
    static String thicknessTypeFromMillimeters(String mmText) {
        String mm = RequestFormTpiPdfFieldLayout.toAsciiDigits(mmText);
        if (mm.contains(".")) {
            return mm.replace(".", "");
        }
        return mm.length() == 1 ? mm + "0" : mm;
    }

    private static void appendHbRows(
            List<HbGbRow> rows, List<String> hinmeis, List<String> specs, List<String> qtys) {
        for (HbGbRow row : rows) {
            String typeCode = thicknessTypeFromMillimeters(row.thicknessMm());
            hinmeis.add(row.part());
            specs.add(
                    JuchuSheetColumnLayout.buildSpecName(
                            row.part(), typeCode, row.width(), row.length()));
            qtys.add(row.rollQty());
        }
    }

    private static void appendHbRawMaterialRows(List<HbGbRow> rows, Map<String, String> raw) {
        if (rows == null || rows.isEmpty()) {
            return;
        }
        List<String> hinmeis = new ArrayList<>();
        List<String> specs = new ArrayList<>();
        List<String> qtys = new ArrayList<>();
        appendHbRows(rows, hinmeis, specs, qtys);
        raw.put("原反品名", String.join("\n", hinmeis));
        raw.put("品名1", String.join("\n", hinmeis));
        raw.put("原反", String.join("\n", specs));
        raw.put("原反数量", String.join("\n", qtys));
    }

    private static void appendGenericRows(
            List<GenericGbRow> rows,
            List<String> hinmeis,
            List<String> specs,
            List<String> qtys,
            List<String> colors) {
        for (GenericGbRow row : rows) {
            hinmeis.add(row.part());
            specs.add(
                    JuchuSheetColumnLayout.buildSpecName(
                            row.part(), typeTokenFromGenericRow(row.typeRaw()), row.width(), row.length()));
            qtys.add(row.rollQty());
            colors.add(colorTokenFromGenericRow(row.typeRaw()));
        }
    }

    private static void appendGenericRawMaterialRows(List<GenericGbRow> rows, Map<String, String> raw) {
        if (rows == null || rows.isEmpty()) {
            return;
        }
        List<String> hinmeis = new ArrayList<>();
        List<String> specs = new ArrayList<>();
        List<String> qtys = new ArrayList<>();
        List<String> colors = new ArrayList<>();
        appendGenericRows(rows, hinmeis, specs, qtys, colors);
        raw.put("原反品名", String.join("\n", hinmeis));
        raw.put("品名1", String.join("\n", hinmeis));
        raw.put("原反", String.join("\n", specs));
        raw.put("原反数量", String.join("\n", qtys));
        if (colors.stream().anyMatch(color -> color != null && !color.isBlank())) {
            raw.put("原反色", String.join("\n", colors));
        }
    }

    /** {@code NR28 BK0} → {@code NR28}、{@code AW07} はそのまま。 */
    static String typeTokenFromGenericRow(String typeRaw) {
        if (typeRaw == null || typeRaw.isBlank()) {
            return "";
        }
        String normalized = typeRaw.strip();
        int space = normalized.indexOf(' ');
        return space > 0 ? normalized.substring(0, space).strip() : normalized;
    }

    /** {@code NR28 BK0} → {@code BK0}。 */
    static String colorTokenFromGenericRow(String typeRaw) {
        if (typeRaw == null || typeRaw.isBlank()) {
            return "";
        }
        String normalized = typeRaw.strip();
        int space = normalized.indexOf(' ');
        return space > 0 ? normalized.substring(space + 1).strip() : "";
    }

    private static String collapseWhitespace(String text) {
        if (text == null || text.isBlank()) {
            return "";
        }
        return text.replaceAll("\\s+", " ").strip();
    }

    private static void fillMeta(Map<String, String> raw, String fileName) {
        raw.put("原本ファイル名", fileName != null ? fileName : "");
        raw.put("原本シート名", RequestFormTpiPdfFieldLayout.META_SHEET_NAME);
        raw.put(
                RequestFormTpiPdfFieldLayout.META_SOURCE_KIND,
                RequestFormTpiPdfFieldLayout.META_SOURCE_KIND_TPI_PDF);
        raw.put(RequestFormTpiPdfFieldLayout.META_TPI_LAYOUT, RequestFormTpiPdfFieldLayout.LAYOUT_GB_SLICE);
    }
}
