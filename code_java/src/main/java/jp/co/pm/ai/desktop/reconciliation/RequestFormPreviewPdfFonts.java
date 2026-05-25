package jp.co.pm.ai.desktop.reconciliation;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

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
            candidates.add(fonts.resolve("msgothic.ttc"));
            candidates.add(fonts.resolve("YuGothM.ttc"));
            candidates.add(fonts.resolve("meiryo.ttc"));
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
        PDFont regular = PDType0Font.load(document, fontFile);
        PDFont bold = regular;
        String lower = fontFile.getName().toLowerCase();
        if (lower.contains("goth") || lower.contains("meiryo")) {
            Path boldCandidate = fontFile.toPath().getParent().resolve("YuGothB.ttc");
            if (Files.isRegularFile(boldCandidate)) {
                try {
                    bold = PDType0Font.load(document, boldCandidate.toFile());
                } catch (IOException ignored) {
                    bold = regular;
                }
            }
        }
        return new FontPair(regular, bold);
    }
}
