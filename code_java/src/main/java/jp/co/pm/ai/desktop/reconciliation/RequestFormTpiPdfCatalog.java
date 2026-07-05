package jp.co.pm.ai.desktop.reconciliation;

import java.io.File;
import java.io.IOException;
import java.text.Normalizer;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;

/** TPI 依頼書 PDF フォルダ内のファイルを依頼Ｎｏへ関連付ける。 */
public final class RequestFormTpiPdfCatalog {

    private RequestFormTpiPdfCatalog() {}

    /**
     * 依頼Ｎｏに対応する TPI PDF を探索する（ファイル名のみ）。
     *
     * <p>完全一致 stem → ファイル名内の依頼Ｎｏ → ファイル名に依頼Ｎｏを含む、の順。
     * 短い stem 前方一致（{@code GB.pdf} → 任意の {@code GB60xxx}）は誤関連付けの原因となるため行わない。
     */
    public static Optional<File> findForIraiNo(String iraiNo, String tpiPdfFolder) {
        if (iraiNo == null || iraiNo.isBlank() || tpiPdfFolder == null || tpiPdfFolder.isBlank()) {
            return Optional.empty();
        }
        File dir = new File(tpiPdfFolder);
        if (!dir.isDirectory()) {
            return Optional.empty();
        }
        File[] pdfs =
                dir.listFiles(
                        (d, name) ->
                                name != null
                                        && name.toLowerCase(Locale.ROOT).endsWith(".pdf")
                                        && !name.startsWith("~$"));
        if (pdfs == null || pdfs.length == 0) {
            return Optional.empty();
        }
        String normIrai = normalizeKey(iraiNo);

        for (File pdf : pdfs) {
            String stem = stemWithoutExtension(pdf.getName());
            if (normalizeKey(stem).equals(normIrai)) {
                return Optional.of(pdf);
            }
        }

        for (File pdf : pdfs) {
            String fromName = RequestFormTpiPdfFieldLayout.parseIraiNoFromFileName(pdf.getName());
            if (!fromName.isBlank() && normalizeKey(fromName).equals(normIrai)) {
                return Optional.of(pdf);
            }
        }

        for (File pdf : pdfs) {
            if (normalizeKey(pdf.getName()).contains(normIrai)) {
                return Optional.of(pdf);
            }
        }

        return Optional.empty();
    }

    /**
     * PDF 本文（テキスト抽出または OCR）の依頼Ｎｏが一致するファイルを探索する。
     * ファイル名が {@code GB.pdf} のように短い場合の正規経路。
     */
    public static Optional<File> findForIraiNoByPdfContent(
            String iraiNo, String tpiPdfFolder, Map<String, String> ui, File parseCacheRoot) {
        if (iraiNo == null || iraiNo.isBlank() || tpiPdfFolder == null || tpiPdfFolder.isBlank()) {
            return Optional.empty();
        }
        File dir = new File(tpiPdfFolder);
        if (!dir.isDirectory()) {
            return Optional.empty();
        }
        File[] pdfs =
                dir.listFiles(
                        (d, name) ->
                                name != null
                                        && name.toLowerCase(Locale.ROOT).endsWith(".pdf")
                                        && !name.startsWith("~$"));
        if (pdfs == null || pdfs.length == 0) {
            return Optional.empty();
        }
        String normIrai = normalizeKey(iraiNo);
        for (File pdf : pdfs) {
            try {
                if (pdfContentMatchesIraiNo(pdf, normIrai, ui, parseCacheRoot)) {
                    return Optional.of(pdf);
                }
            } catch (Exception ex) {
                System.err.println(
                        "TPI PDF 本文照合失敗 " + pdf.getName() + ": " + ex.getMessage());
            }
        }
        return Optional.empty();
    }

    private static boolean pdfContentMatchesIraiNo(
            File pdf, String normIrai, Map<String, String> ui, File parseCacheRoot)
            throws IOException {
        Optional<List<Map<String, String>>> cached =
                parseCacheRoot != null
                        ? RequestFormSourceCache.loadParseEntries(parseCacheRoot, pdf)
                        : Optional.empty();
        List<Map<String, String>> parsed;
        if (cached.isPresent() && !cached.get().isEmpty()) {
            parsed =
                    RequestFormTpiPdfSplitter.ensureSplitPdfs(
                            pdf, cached.get(), parseCacheRoot, ui);
            if (parseCacheRoot != null && !parsed.isEmpty()) {
                RequestFormSourceCache.saveParseEntries(parseCacheRoot, pdf, parsed);
            }
        } else {
            parsed = RequestFormTpiPdfExtractor.extractEntriesWithSplit(pdf, ui, parseCacheRoot);
            if (parseCacheRoot != null && !parsed.isEmpty()) {
                RequestFormSourceCache.saveParseEntries(parseCacheRoot, pdf, parsed);
            }
        }
        for (Map<String, String> entry : parsed) {
            String contentIrai = entry.get("依頼Ｎｏ");
            if (contentIrai != null
                    && !contentIrai.isBlank()
                    && normalizeKey(contentIrai).equals(normIrai)) {
                return true;
            }
        }
        return false;
    }

    /**
     * 複数依頼が同一 PDF に束ねられている場合の自動追加可否。
     * 専用 PDF（例: {@code GB60604.pdf}）が別途ある依頼は束ね PDF 側を追加しない。
     */
    static boolean shouldAutoAddScannedEntry(
            File tpiDir, File pdf, List<Map<String, String>> entries, Map<String, String> entry) {
        if (entry == null) {
            return false;
        }
        String iraiNo = entry.get("依頼Ｎｏ");
        if (iraiNo == null || iraiNo.isBlank()) {
            return false;
        }
        if (entries == null || entries.size() <= 1) {
            return true;
        }
        if (hasDedicatedPdfFile(tpiDir, iraiNo)) {
            return false;
        }
        return true;
    }

    /** 共有 PDF 内の依頼が受注行へ関連付け可能か（分割後は各依頼を独立扱い）。 */
    static boolean canLinkIraiInSharedPdf(
            String iraiNo, File pdf, List<Map<String, String>> entries, File tpiDir) {
        if (iraiNo == null
                || iraiNo.isBlank()
                || entries == null
                || entries.isEmpty()
                || pdf == null) {
            return false;
        }
        for (Map<String, String> entry : entries) {
            if (!entryMatchesIrai(entry, iraiNo)) {
                continue;
            }
            if (hasDedicatedPdfFile(tpiDir, iraiNo)
                    && !dedicatedPdfFile(tpiDir, iraiNo).equals(pdf)) {
                return false;
            }
            return true;
        }
        return false;
    }

    static boolean hasDedicatedPdfFile(File tpiDir, String iraiNo) {
        File dedicated = dedicatedPdfFile(tpiDir, iraiNo);
        return dedicated != null && dedicated.isFile();
    }

    private static File dedicatedPdfFile(File tpiDir, String iraiNo) {
        if (tpiDir == null || iraiNo == null || iraiNo.isBlank()) {
            return null;
        }
        return new File(tpiDir, iraiNo + ".pdf");
    }

    private static boolean entryMatchesIrai(Map<String, String> entry, String iraiNo) {
        if (entry == null || iraiNo == null) {
            return false;
        }
        String parsed = entry.get("依頼Ｎｏ");
        return parsed != null && normalizeKey(parsed).equals(normalizeKey(iraiNo));
    }

    static String extractIraiNoFromPdf(File pdf, Map<String, String> ui, File parseCacheRoot)
            throws IOException {
        Optional<List<Map<String, String>>> cached =
                parseCacheRoot != null
                        ? RequestFormSourceCache.loadParseEntries(parseCacheRoot, pdf)
                        : Optional.empty();
        if (cached.isPresent() && !cached.get().isEmpty()) {
            for (Map<String, String> entry : cached.get()) {
                String fromCache = entry.get("依頼Ｎｏ");
                if (fromCache != null && !fromCache.isBlank()) {
                    return fromCache;
                }
            }
        }
        List<Map<String, String>> parsed =
                RequestFormTpiPdfExtractor.extractEntriesWithSplit(pdf, ui, parseCacheRoot);
        if (parseCacheRoot != null && !parsed.isEmpty()) {
            RequestFormSourceCache.saveParseEntries(parseCacheRoot, pdf, parsed);
        }
        for (Map<String, String> entry : parsed) {
            String irai = entry.get("依頼Ｎｏ");
            if (irai != null && !irai.isBlank()) {
                return irai;
            }
        }
        return "";
    }

    /** プレビュー・照合の最低限 rawMap（本文照合不可時は関連付けに使わない）。 */
    public static Map<String, String> buildMinimalRaw(File pdf, String iraiNo) {
        Map<String, String> raw = new LinkedHashMap<>();
        raw.put("依頼Ｎｏ", iraiNo != null ? iraiNo.strip() : "");
        raw.put("原本ファイル名", pdf != null ? pdf.getName() : "");
        raw.put("原本シート名", RequestFormTpiPdfFieldLayout.META_SHEET_NAME);
        raw.put(
                RequestFormTpiPdfFieldLayout.META_SOURCE_KIND,
                RequestFormTpiPdfFieldLayout.META_SOURCE_KIND_TPI_PDF);
        raw.put(RequestFormTpiPdfFieldLayout.META_READ_MODE, "FILENAME_LINK");
        return raw;
    }

    static String stemWithoutExtension(String fileName) {
        if (fileName == null || fileName.isBlank()) {
            return "";
        }
        int dot = fileName.lastIndexOf('.');
        return dot > 0 ? fileName.substring(0, dot) : fileName;
    }

    static String normalizeKey(String val) {
        if (val == null) {
            return "";
        }
        String text = val.strip().toUpperCase(Locale.ROOT);
        text = Normalizer.normalize(text, Normalizer.Form.NFKC);
        text = text.replaceAll("\\s+", "");
        text = text.replace("－", "-").replace("ー", "-").replace("―", "-").replace("‐", "-");
        return text;
    }
}
