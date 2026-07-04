package jp.co.pm.ai.desktop.reconciliation;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

/** 1 つの TPI PDF に複数依頼が束ねられている場合、依頼単位 PDF へ分割する。 */
final class RequestFormTpiPdfSplitter {

    private static final Pattern GB_IRAI_TOKEN =
            Pattern.compile("G\\s*B\\s*([\\d０-９]{4,6})", Pattern.CASE_INSENSITIVE);

    private RequestFormTpiPdfSplitter() {}

    static List<Map<String, String>> attachSplitPdfsIfBundled(
            File sourcePdf,
            List<Map<String, String>> entries,
            File cacheRoot,
            Map<String, String> ui)
            throws IOException {
        if (sourcePdf == null
                || !sourcePdf.isFile()
                || entries == null
                || entries.isEmpty()
                || cacheRoot == null) {
            return entries != null ? entries : List.of();
        }
        if (entries.size() <= 1) {
            return entries;
        }
        String layout = entries.get(0).get(RequestFormTpiPdfFieldLayout.META_TPI_LAYOUT);
        if (!RequestFormTpiPdfFieldLayout.LAYOUT_GB_SLICE.equals(layout)) {
            return entries;
        }
        try (PDDocument doc = Loader.loadPDF(sourcePdf)) {
            int pageCount = doc.getNumberOfPages();
            if (pageCount <= 0) {
                return entries;
            }
            PDFTextStripper stripper = new PDFTextStripper();
            stripper.setSortByPosition(true);
            List<Integer> startPages = resolveStartPages(doc, stripper, entries);
            List<Map<String, String>> out = new ArrayList<>(entries.size());
            for (int i = 0; i < entries.size(); i++) {
                int startPage0 = startPages.get(i);
                int endPage0 =
                        i + 1 < startPages.size()
                                ? Math.max(startPage0, startPages.get(i + 1) - 1)
                                : pageCount - 1;
                Map<String, String> entry = entries.get(i);
                String iraiNo = entry.get("依頼Ｎｏ");
                File splitFile =
                        RequestFormSourceCache.splitCacheFile(cacheRoot, sourcePdf, iraiNo);
                if (!RequestFormSourceCache.isSplitCacheValid(
                        splitFile, sourcePdf, startPage0, endPage0)) {
                    writePageRange(doc, startPage0, endPage0, splitFile);
                    RequestFormSourceCache.writeSplitCacheMeta(
                            splitFile, sourcePdf, startPage0, endPage0);
                }
                Map<String, String> copy = new LinkedHashMap<>(entry);
                copy.put(
                        RequestFormTpiPdfFieldLayout.META_SPLIT_PDF_PATH,
                        splitFile.getAbsolutePath());
                out.add(copy);
            }
            return out;
        }
    }

    /** 分割 PDF が欠落している parse キャッシュを再分割する。 */
    static List<Map<String, String>> ensureSplitPdfs(
            File sourcePdf,
            List<Map<String, String>> entries,
            File cacheRoot,
            Map<String, String> ui)
            throws IOException {
        if (entries == null || entries.size() <= 1 || cacheRoot == null) {
            return entries != null ? entries : List.of();
        }
        boolean missingSplit =
                entries.stream()
                        .anyMatch(
                                entry -> {
                                    String path =
                                            entry.get(
                                                    RequestFormTpiPdfFieldLayout
                                                            .META_SPLIT_PDF_PATH);
                                    return path == null
                                            || path.isBlank()
                                            || !new File(path).isFile();
                                });
        if (!missingSplit) {
            return entries;
        }
        return attachSplitPdfsIfBundled(sourcePdf, entries, cacheRoot, ui);
    }

    private static List<Integer> resolveStartPages(
            PDDocument doc, PDFTextStripper stripper, List<Map<String, String>> entries)
            throws IOException {
        List<Integer> starts = new ArrayList<>(entries.size());
        int searchFrom = 0;
        for (Map<String, String> entry : entries) {
            int page0 = findIraiStartPage(doc, stripper, entry.get("依頼Ｎｏ"), searchFrom);
            starts.add(page0);
            searchFrom = Math.min(page0 + 1, doc.getNumberOfPages() - 1);
        }
        return starts;
    }

    private static int findIraiStartPage(
            PDDocument doc, PDFTextStripper stripper, String iraiNo, int fromPage0)
            throws IOException {
        if (iraiNo == null || iraiNo.isBlank()) {
            return Math.max(0, Math.min(fromPage0, doc.getNumberOfPages() - 1));
        }
        String normIrai = RequestFormTpiPdfCatalog.normalizeKey(iraiNo);
        int last = doc.getNumberOfPages() - 1;
        for (int page0 = Math.max(0, fromPage0); page0 <= last; page0++) {
            stripper.setStartPage(page0 + 1);
            stripper.setEndPage(page0 + 1);
            String pageText = stripper.getText(doc);
            if (pageContainsIrai(pageText, normIrai)) {
                return page0;
            }
        }
        return Math.max(0, Math.min(fromPage0, last));
    }

    private static boolean pageContainsIrai(String pageText, String normIrai) {
        if (pageText == null || pageText.isBlank() || normIrai == null || normIrai.isBlank()) {
            return false;
        }
        String normPage = RequestFormTpiPdfCatalog.normalizeKey(pageText);
        if (normPage.contains(normIrai)) {
            return true;
        }
        Matcher token = GB_IRAI_TOKEN.matcher(RequestFormTpiPdfFieldLayout.normalizeText(pageText));
        while (token.find()) {
            String digits = RequestFormTpiPdfFieldLayout.toAsciiDigits(token.group(1));
            String candidate = RequestFormTpiPdfCatalog.normalizeKey("GB" + digits);
            if (candidate.equals(normIrai)) {
                return true;
            }
        }
        return false;
    }

    private static void writePageRange(
            PDDocument source, int startPage0, int endPage0, File out) throws IOException {
        Files.createDirectories(out.getParentFile().toPath());
        try (PDDocument split = new PDDocument()) {
            for (int page0 = startPage0; page0 <= endPage0; page0++) {
                split.importPage(source.getPage(page0));
            }
            split.save(out);
        }
    }
}
