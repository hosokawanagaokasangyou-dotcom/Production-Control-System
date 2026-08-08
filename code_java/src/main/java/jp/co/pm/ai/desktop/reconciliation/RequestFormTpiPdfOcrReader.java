package jp.co.pm.ai.desktop.reconciliation;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.AppPaths.TesseractConfig;
import net.sourceforge.tess4j.Tesseract;
import net.sourceforge.tess4j.TesseractException;

/** 画像スキャン TPI 依頼書 PDF を OCR でテキスト化する。 */
final class RequestFormTpiPdfOcrReader {

    private static final float OCR_RENDER_DPI = 300f;

    private RequestFormTpiPdfOcrReader() {}

    static boolean isAvailable(Map<String, String> ui) {
        return AppPaths.resolveTesseractConfig(ui).isPresent();
    }

    static String readPdfText(File pdfFile, Map<String, String> ui) throws IOException {
        TesseractConfig config =
                AppPaths.resolveTesseractConfig(ui)
                        .orElseThrow(
                                () ->
                                        new IOException(
                                                AppPaths.explainTesseractConfigMissing(ui)));
        try (PDDocument document = Loader.loadPDF(pdfFile)) {
            PDFRenderer renderer = new PDFRenderer(document);
            Tesseract tesseract = createTesseract(config);
            StringBuilder sb = new StringBuilder();
            int pages = document.getNumberOfPages();
            for (int pageIndex = 0; pageIndex < pages; pageIndex++) {
                BufferedImage image =
                        renderer.renderImageWithDPI(pageIndex, OCR_RENDER_DPI, ImageType.RGB);
                try {
                    sb.append(tesseract.doOCR(image));
                } catch (TesseractException ex) {
                    throw new IOException(
                            "TPI PDF OCR 失敗: " + pdfFile.getName() + " (page " + (pageIndex + 1) + ")", ex);
                }
                if (pageIndex + 1 < pages) {
                    sb.append('\n');
                }
            }
            return sb.toString();
        }
    }

    private static Tesseract createTesseract(TesseractConfig config) {
        Tesseract tesseract = new Tesseract();
        Path tessData = config.tessDataDir();
        if (tessData != null && Files.isDirectory(tessData)) {
            String dataPath =
                    tessData.toAbsolutePath().normalize().toString().replace('\\', '/');
            tesseract.setDatapath(dataPath);
            System.setProperty("TESSDATA_PREFIX", dataPath.endsWith("/") ? dataPath : dataPath + "/");
        }
        Path executable = config.executable();
        if (executable != null && Files.isRegularFile(executable)) {
            System.setProperty("net.sourceforge.tess4j.extractor.TesseractExtractor.tesseractPath", executable.toString());
        }
        tesseract.setLanguage("jpn+eng");
        // 1 = OSD 必須（osd.traineddata）。スキャン PDF 全文 OCR は 3（自動分割・OSD 不要）。
        tesseract.setPageSegMode(3);
        tesseract.setOcrEngineMode(1);
        return tesseract;
    }
}
