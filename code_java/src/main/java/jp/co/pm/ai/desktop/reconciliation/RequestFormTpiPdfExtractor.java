package jp.co.pm.ai.desktop.reconciliation;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

import jp.co.pm.ai.desktop.reconciliation.RequestFormTpiPdfFieldLayout.LayoutKind;

/** TPI 依頼書 PDF（QR-06-011）から受注照合用 rawMap を抽出する。 */
public final class RequestFormTpiPdfExtractor {

    private RequestFormTpiPdfExtractor() {}

    /**
     * PDF 1 ファイルから依頼 1 件分の rawMap を返す（現行様式は 1 PDF = 1 依頼）。
     *
     * @throws IOException PDF 読込失敗
     */
    public static List<Map<String, String>> extractEntries(File pdfFile) throws IOException {
        return extractEntries(pdfFile, Map.of());
    }

    /**
     * PDF 1 ファイルから依頼 1 件分の rawMap を返す。{@code ui} は Tesseract 設定解決に使う（空可）。
     *
     * @throws IOException PDF 読込失敗
     */
    public static List<Map<String, String>> extractEntries(File pdfFile, Map<String, String> ui)
            throws IOException {
        return extractEntries(pdfFile, ui, null);
    }

    /**
     * @param ocrPageProgress 画像 PDF の OCR 中にページ進捗を通知する（{@code null} 可）。
     */
    public static List<Map<String, String>> extractEntries(
            File pdfFile, Map<String, String> ui, Consumer<String> ocrPageProgress)
            throws IOException {
        if (pdfFile == null || !pdfFile.isFile()) {
            return List.of();
        }
        RequestFormTpiPdfContentKind contentKind = RequestFormTpiPdfContentDetector.detect(pdfFile);
        String text =
                contentKind == RequestFormTpiPdfContentKind.TEXT
                        ? readPdfText(pdfFile)
                        : RequestFormTpiPdfOcrReader.readPdfText(pdfFile, ui, ocrPageProgress);
        String readMode =
                contentKind == RequestFormTpiPdfContentKind.TEXT
                        ? RequestFormTpiPdfFieldLayout.READ_MODE_TEXT
                        : RequestFormTpiPdfFieldLayout.READ_MODE_OCR;
        List<Map<String, String>> entries = buildAllEntries(pdfFile.getName(), text, readMode);
        return entries;
    }

    /**
     * PDF を解析し、複数依頼が束ねられている場合は依頼単位 PDF へ分割してから返す。
     */
    static List<Map<String, String>> extractEntriesWithSplit(
            File pdfFile, Map<String, String> ui, File parseCacheRoot) throws IOException {
        return extractEntriesWithSplit(pdfFile, ui, parseCacheRoot, null);
    }

    static List<Map<String, String>> extractEntriesWithSplit(
            File pdfFile,
            Map<String, String> ui,
            File parseCacheRoot,
            Consumer<String> ocrPageProgress)
            throws IOException {
        List<Map<String, String>> entries = extractEntries(pdfFile, ui, ocrPageProgress);
        if (parseCacheRoot == null || entries.size() <= 1) {
            return entries;
        }
        return RequestFormTpiPdfSplitter.attachSplitPdfsIfBundled(
                pdfFile, entries, parseCacheRoot, ui);
    }

    /** テスト・デバッグ向け: プレーンテキストから rawMap を組み立てる（先頭依頼）。 */
    static Map<String, String> extractFromTextForTest(String fileName, String text) {
        List<Map<String, String>> entries =
                buildAllEntries(fileName, text, RequestFormTpiPdfFieldLayout.READ_MODE_TEXT);
        return entries.isEmpty() ? Map.of() : entries.get(0);
    }

    private static List<Map<String, String>> buildAllEntries(
            String fileName, String text, String readMode) {
        LayoutKind kind = RequestFormTpiPdfFieldLayout.detectLayout(fileName, text);
        if (kind == LayoutKind.GB_SLICE) {
            List<Map<String, String>> gbEntries =
                    RequestFormTpiPdfLayoutGb.buildAllRawMaps(fileName, text);
            if (gbEntries.isEmpty()) {
                gbEntries = List.of(RequestFormTpiPdfLayoutGb.buildRawMap(fileName, text));
            }
            return finalizeEntries(fileName, text, readMode, gbEntries);
        }
        Map<String, String> raw =
                switch (kind) {
                    case ECOWD -> RequestFormTpiPdfLayoutEcowd.buildRawMap(fileName, text);
                    case GB_SLICE -> RequestFormTpiPdfLayoutGb.buildRawMap(fileName, text);
                    case PN -> RequestFormTpiPdfLayoutPn.buildRawMap(fileName, text);
                };
        return List.of(finalizeEntry(fileName, text, readMode, raw));
    }

    private static List<Map<String, String>> finalizeEntries(
            String fileName,
            String text,
            String readMode,
            List<Map<String, String>> entries) {
        List<Map<String, String>> out = new ArrayList<>(entries.size());
        for (Map<String, String> entry : entries) {
            out.add(finalizeEntry(fileName, text, readMode, entry));
        }
        return out;
    }

    private static Map<String, String> finalizeEntry(
            String fileName, String text, String readMode, Map<String, String> raw) {
        Map<String, String> resolved = new LinkedHashMap<>(raw);
        if (resolved.get("依頼Ｎｏ") == null || resolved.get("依頼Ｎｏ").isBlank()) {
            resolved.put("依頼Ｎｏ", RequestFormTpiPdfFieldLayout.parseIraiNoFromFileName(fileName));
            if (resolved.get("依頼Ｎｏ").isBlank()) {
                resolved.put("依頼Ｎｏ", RequestFormTpiPdfFieldLayout.parseIraiNoFromText(text));
            }
        }
        resolved.put(RequestFormTpiPdfFieldLayout.META_READ_MODE, readMode);
        return resolved;
    }

    private static Map<String, String> buildRawMap(String fileName, String text, String readMode) {
        List<Map<String, String>> entries = buildAllEntries(fileName, text, readMode);
        return entries.isEmpty() ? Map.of() : entries.get(0);
    }

    static String readPdfText(File pdfFile) throws IOException {
        try (PDDocument document = Loader.loadPDF(pdfFile)) {
            PDFTextStripper stripper = new PDFTextStripper();
            stripper.setSortByPosition(true);
            return stripper.getText(document);
        }
    }
}
