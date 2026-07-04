package jp.co.pm.ai.desktop.reconciliation;

import java.io.File;
import java.io.IOException;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDResources;
import org.apache.pdfbox.pdmodel.graphics.PDXObject;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.text.PDFTextStripper;

/** TPI 依頼書 PDF がテキスト PDF か画像スキャン PDF かを自動判定する。 */
final class RequestFormTpiPdfContentDetector {

    /** テキスト PDF とみなす有意文字数の下限。 */
    private static final int MIN_MEANINGFUL_CHARS_FOR_TEXT = 50;

    private RequestFormTpiPdfContentDetector() {}

    static RequestFormTpiPdfContentKind detect(File pdfFile) throws IOException {
        if (pdfFile == null || !pdfFile.isFile()) {
            return RequestFormTpiPdfContentKind.IMAGE_SCAN;
        }
        try (PDDocument document = Loader.loadPDF(pdfFile)) {
            PDFTextStripper stripper = new PDFTextStripper();
            stripper.setSortByPosition(true);
            String text = stripper.getText(document);
            return detectFromExtractedText(document, text);
        }
    }

    static RequestFormTpiPdfContentKind detectFromExtractedText(PDDocument document, String text)
            throws IOException {
        String normalized = RequestFormTpiPdfFieldLayout.normalizeText(text);
        int meaningful = countMeaningfulChars(normalized);
        boolean markers = hasTpiRequestFormMarkers(normalized);
        if (meaningful >= MIN_MEANINGFUL_CHARS_FOR_TEXT && markers) {
            return RequestFormTpiPdfContentKind.TEXT;
        }
        if (document != null && isImageDominantDocument(document, meaningful)) {
            return RequestFormTpiPdfContentKind.IMAGE_SCAN;
        }
        if (meaningful < MIN_MEANINGFUL_CHARS_FOR_TEXT || !markers) {
            return RequestFormTpiPdfContentKind.IMAGE_SCAN;
        }
        return RequestFormTpiPdfContentKind.TEXT;
    }

    private static boolean hasTpiRequestFormMarkers(String normalized) {
        if (normalized.isBlank()) {
            return false;
        }
        if (!RequestFormTpiPdfFieldLayout.parseIraiNoFromText(normalized).isBlank()) {
            return true;
        }
        return normalized.contains("注文書")
                || normalized.contains("加工製品")
                || normalized.contains("投入原反")
                || normalized.contains("希望納期")
                || normalized.contains("古河原反")
                || normalized.contains("スライス")
                || normalized.contains("依頼No")
                || normalized.contains("依頼Ｎｏ")
                || normalized.contains("東レ")
                || normalized.contains("QR")
                || normalized.contains("ＱＲ")
                || normalized.contains("ECOWD")
                || normalized.contains("後加工");
    }

    private static int countMeaningfulChars(String text) {
        if (text == null || text.isBlank()) {
            return 0;
        }
        int count = 0;
        for (int i = 0; i < text.length(); i++) {
            char c = text.charAt(i);
            if (Character.isLetterOrDigit(c) || isCjk(c)) {
                count++;
            }
        }
        return count;
    }

    private static boolean isCjk(char c) {
        Character.UnicodeBlock block = Character.UnicodeBlock.of(c);
        return block == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS
                || block == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS_EXTENSION_A
                || block == Character.UnicodeBlock.HIRAGANA
                || block == Character.UnicodeBlock.KATAKANA
                || block == Character.UnicodeBlock.HALFWIDTH_AND_FULLWIDTH_FORMS;
    }

    private static boolean isImageDominantDocument(PDDocument document, int meaningfulTextChars)
            throws IOException {
        if (document == null || document.getNumberOfPages() <= 0) {
            return meaningfulTextChars < MIN_MEANINGFUL_CHARS_FOR_TEXT;
        }
        int imagePages = 0;
        int pages = document.getNumberOfPages();
        for (int i = 0; i < pages; i++) {
            PDPage page = document.getPage(i);
            if (pageHasLargeImage(page)) {
                imagePages++;
            }
        }
        return imagePages > 0 && meaningfulTextChars < MIN_MEANINGFUL_CHARS_FOR_TEXT;
    }

    private static boolean pageHasLargeImage(PDPage page) throws IOException {
        PDResources resources = page.getResources();
        if (resources == null) {
            return false;
        }
        for (org.apache.pdfbox.cos.COSName name : resources.getXObjectNames()) {
            PDXObject xObject = resources.getXObject(name);
            if (xObject instanceof PDImageXObject image) {
                if (image.getWidth() >= 400 && image.getHeight() >= 400) {
                    return true;
                }
            }
        }
        return false;
    }
}
