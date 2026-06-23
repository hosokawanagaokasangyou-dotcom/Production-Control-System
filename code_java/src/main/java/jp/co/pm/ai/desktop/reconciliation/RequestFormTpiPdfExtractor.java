package jp.co.pm.ai.desktop.reconciliation;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

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
        if (pdfFile == null || !pdfFile.isFile()) {
            return List.of();
        }
        String text = readPdfText(pdfFile);
        String fileName = pdfFile.getName();
        LayoutKind kind = RequestFormTpiPdfFieldLayout.detectLayout(fileName, text);
        Map<String, String> raw =
                kind == LayoutKind.ECOWD
                        ? RequestFormTpiPdfLayoutEcowd.buildRawMap(fileName, text)
                        : RequestFormTpiPdfLayoutPn.buildRawMap(fileName, text);
        if (raw.get("依頼Ｎｏ") == null || raw.get("依頼Ｎｏ").isBlank()) {
            Map<String, String> copy = new LinkedHashMap<>(raw);
            copy.put("依頼Ｎｏ", RequestFormTpiPdfFieldLayout.parseIraiNoFromFileName(fileName));
            raw = copy;
        }
        return List.of(raw);
    }

    /** テスト・デバッグ向け: プレーンテキストから rawMap を組み立てる。 */
    static Map<String, String> extractFromTextForTest(String fileName, String text) {
        LayoutKind kind = RequestFormTpiPdfFieldLayout.detectLayout(fileName, text);
        return kind == LayoutKind.ECOWD
                ? RequestFormTpiPdfLayoutEcowd.buildRawMap(fileName, text)
                : RequestFormTpiPdfLayoutPn.buildRawMap(fileName, text);
    }

    static String readPdfText(File pdfFile) throws IOException {
        try (PDDocument document = Loader.loadPDF(pdfFile)) {
            PDFTextStripper stripper = new PDFTextStripper();
            stripper.setSortByPosition(true);
            return stripper.getText(document);
        }
    }
}
