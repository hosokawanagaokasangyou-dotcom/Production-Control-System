package jp.co.pm.ai.desktop.print;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import org.apache.fontbox.ttf.TrueTypeCollection;
import org.apache.fontbox.ttf.TrueTypeFont;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDType0Font;

/** PDFBox 向け CJK フォント解決（依頼書 PDF と同系の OS フォント候補）。 */
final class PdfBoxCjkFontLoader {

    private PdfBoxCjkFontLoader() {}

    static PDFont loadRegular(PDDocument document) throws IOException {
        List<Path> candidates = new ArrayList<>();
        candidates.add(Path.of("/mnt/c/Windows/Fonts/msgothic.ttc"));
        candidates.add(Path.of("/mnt/c/Windows/Fonts/msgothic.ttf"));
        candidates.add(Path.of("/mnt/c/Windows/Fonts/meiryo.ttc"));
        String windir = System.getenv("WINDIR");
        if (windir != null && !windir.isBlank()) {
            Path fonts = Path.of(windir, "Fonts");
            candidates.add(fonts.resolve("msgothic.ttf"));
            candidates.add(fonts.resolve("msgothic.ttc"));
            candidates.add(fonts.resolve("meiryo.ttc"));
            candidates.add(fonts.resolve("YuGothM.ttc"));
        }
        candidates.add(Path.of("/usr/share/fonts/truetype/fonts-japanese-gothic.ttf"));

        IOException lastError = null;
        for (Path candidate : candidates) {
            if (!Files.isRegularFile(candidate)) {
                continue;
            }
            try {
                return loadFromFile(document, candidate.toFile());
            } catch (IOException ex) {
                lastError = ex;
            }
        }
        if (lastError != null) {
            throw lastError;
        }
        throw new IOException(
                "PDF 用の日本語フォントが見つかりません（Windows Fonts または Noto CJK を確認してください）");
    }

    private static PDFont loadFromFile(PDDocument document, File fontFile) throws IOException {
        String lower = fontFile.getName().toLowerCase(Locale.ROOT);
        if (lower.endsWith(".ttc")) {
            return loadType0FromTtc(document, fontFile, preferredTtcFaceNames(lower));
        }
        return PDType0Font.load(document, fontFile);
    }

    private static PDFont loadType0FromTtc(
            PDDocument document, File ttcFile, List<String> preferredFaceNames) throws IOException {
        try (TrueTypeCollection collection = new TrueTypeCollection(ttcFile)) {
            TrueTypeFont face = resolveTtcFace(collection, preferredFaceNames);
            if (face == null) {
                throw new IOException("TTC に利用可能なフォントがありません: " + ttcFile.getName());
            }
            return PDType0Font.load(document, face, true);
        }
    }

    private static TrueTypeFont resolveTtcFace(
            TrueTypeCollection collection, List<String> preferredFaceNames) throws IOException {
        for (String name : preferredFaceNames) {
            TrueTypeFont font = collection.getFontByName(name);
            if (font != null) {
                return font;
            }
        }
        final TrueTypeFont[] first = new TrueTypeFont[1];
        collection.processAllFonts(
                ttf -> {
                    if (first[0] == null) {
                        first[0] = ttf;
                    }
                });
        return first[0];
    }

    private static List<String> preferredTtcFaceNames(String lowerFileName) {
        if (lowerFileName.contains("msgothic")) {
            return List.of("MS-Gothic", "MS Gothic", "MSゴシック", "ＭＳ ゴシック");
        }
        if (lowerFileName.contains("meiryo")) {
            return List.of("Meiryo", "Meiryo Regular", "メイリオ");
        }
        if (lowerFileName.contains("yugoth")) {
            return List.of("Yu Gothic Medium", "YuGothic-Medium", "Yu Gothic", "游ゴシック Medium");
        }
        if (lowerFileName.contains("noto")) {
            return List.of(
                    "Noto Sans CJK JP",
                    "Noto Sans CJK JP Regular",
                    "NotoSansCJKjp-Regular",
                    "Noto Sans CJK Regular");
        }
        return List.of();
    }
}
