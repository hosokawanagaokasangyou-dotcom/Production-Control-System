package jp.co.pm.ai.desktop.reconciliation;

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

/** 依頼書 PDF プレビュー用の日本語フォント解決。 */
final class RequestFormPreviewPdfFonts {

    private RequestFormPreviewPdfFonts() {}

    record FontPair(PDFont regular, PDFont bold) {}

    static FontPair load(PDDocument document) throws IOException {
        List<Path> candidates = new ArrayList<>();
        String windir = System.getenv("WINDIR");
        if (windir != null && !windir.isBlank()) {
            Path fonts = Path.of(windir, "Fonts");
            candidates.add(fonts.resolve("msgothic.ttf"));
            candidates.add(fonts.resolve("msgothic.ttc"));
            candidates.add(fonts.resolve("meiryo.ttc"));
            candidates.add(fonts.resolve("YuGothM.ttc"));
            candidates.add(fonts.resolve("YuGothB.ttc"));
        }
        candidates.add(Path.of("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"));
        candidates.add(Path.of("/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"));
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
                "依頼書 PDF 用の日本語フォントが見つかりません（Windows Fonts または Noto CJK を確認してください）");
    }

    private static FontPair loadFromFile(PDDocument document, File fontFile) throws IOException {
        String lower = fontFile.getName().toLowerCase(Locale.ROOT);
        PDFont regular;
        if (lower.endsWith(".ttc")) {
            regular = loadType0FromTtc(document, fontFile, preferredTtcFaceNames(lower));
        } else {
            regular = PDType0Font.load(document, fontFile);
        }
        PDFont bold = regular;
        if (lower.contains("goth") || lower.contains("meiryo")) {
            Path boldCandidate = fontFile.toPath().getParent().resolve("YuGothB.ttc");
            if (Files.isRegularFile(boldCandidate)) {
                try {
                    bold =
                            loadType0FromTtc(
                                    document,
                                    boldCandidate.toFile(),
                                    List.of("Yu Gothic Bold", "YuGothic-Bold", "Yu Gothic"));
                } catch (IOException ignored) {
                    bold = regular;
                }
            }
        }
        return new FontPair(regular, bold);
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

    private static List<String> preferredTtcFaceNames(String fileNameLower) {
        if (fileNameLower.contains("msgothic")) {
            return List.of("MS-Gothic", "MS Gothic", "MSゴシック", "ＭＳ ゴシック");
        }
        if (fileNameLower.contains("meiryo")) {
            return List.of("Meiryo", "Meiryo Regular", "メイリオ");
        }
        if (fileNameLower.contains("yugothb")) {
            return List.of("Yu Gothic Bold", "YuGothic-Bold", "Yu Gothic");
        }
        if (fileNameLower.contains("yugoth")) {
            return List.of("Yu Gothic Medium", "YuGothic-Medium", "Yu Gothic", "游ゴシック Medium");
        }
        if (fileNameLower.contains("noto")) {
            return List.of(
                    "Noto Sans CJK JP",
                    "Noto Sans CJK JP Regular",
                    "NotoSansCJKjp-Regular",
                    "Noto Sans CJK Regular");
        }
        return List.of();
    }
}
