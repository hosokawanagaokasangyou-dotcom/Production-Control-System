package jp.co.pm.ai.desktop.print;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDFont;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;

/** 工場別ユーザー管理情報を PDF 化する（管理者向け）。 */
public final class FactoryOperatorUserPdfExporter {

    private static final DateTimeFormatter TS =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss").withZone(ZoneId.systemDefault());

    private static final float MARGIN = 40f;
    private static final float TITLE_SIZE = 14f;
    private static final float BODY_SIZE = 10f;
    private static final float LINE_HEIGHT = 16f;

    public record Row(String name, String pinStatus, String adminPin) {}

    private FactoryOperatorUserPdfExporter() {}

    public static Path resolveOutputPath(Map<String, String> ui, FactorySite site) {
        if (AppPaths.usesRemoteDesktopAppHome()) {
            return AppPaths.resolveRdpLauncherOperatorUsersPdfPath(ui);
        }
        return AppPaths.factoryOperatorUsersPdfPath(ui, site);
    }

    /**
     * 工場別ユーザー一覧を PDF へ原子的に書き出す。
     *
     * @param storePathLabel データ正本パス（表示用。null 可）
     */
    public static void export(
            Path outputPath,
            FactorySite site,
            List<Row> rows,
            String exportedBy,
            Instant exportedAt,
            String storePathLabel)
            throws IOException {
        Objects.requireNonNull(outputPath, "outputPath");
        Objects.requireNonNull(site, "site");
        Objects.requireNonNull(rows, "rows");
        Instant when = exportedAt != null ? exportedAt : Instant.now();
        Path parent = outputPath.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        Path temp =
                Files.createTempFile(
                        parent != null ? parent : Path.of("."),
                        "factory-operator-users-",
                        ".pdf.tmp");
        try {
            writePdf(temp, site, rows, exportedBy, when, storePathLabel);
            Files.move(
                    temp,
                    outputPath,
                    StandardCopyOption.REPLACE_EXISTING,
                    StandardCopyOption.ATOMIC_MOVE);
        } catch (IOException | RuntimeException ex) {
            try {
                Files.deleteIfExists(temp);
            } catch (IOException ignored) {
                // ignore cleanup failure
            }
            throw ex instanceof IOException ioe ? ioe : new IOException(ex);
        }
    }

    private static void writePdf(
            Path path,
            FactorySite site,
            List<Row> rows,
            String exportedBy,
            Instant exportedAt,
            String storePathLabel)
            throws IOException {
        try (PDDocument document = new PDDocument()) {
            PDFont font = PdfBoxCjkFontLoader.loadRegular(document);
            PDPage page = new PDPage(PDRectangle.A4);
            document.addPage(page);
            float y = page.getMediaBox().getHeight() - MARGIN;
            try (PDPageContentStream cs =
                    new PDPageContentStream(document, page, PDPageContentStream.AppendMode.OVERWRITE, true, true)) {
                y = drawLine(cs, font, TITLE_SIZE, MARGIN, y, "ユーザー管理情報（管理者）");
                y -= LINE_HEIGHT * 0.5f;
                y = drawLine(cs, font, BODY_SIZE, MARGIN, y, "工場: " + site.displayLabelJa());
                y = drawLine(cs, font, BODY_SIZE, MARGIN, y, "出力日時: " + TS.format(exportedAt));
                String by =
                        exportedBy != null && !exportedBy.isBlank() ? exportedBy.strip() : "（不明）";
                y = drawLine(cs, font, BODY_SIZE, MARGIN, y, "出力者: " + by);
                if (storePathLabel != null && !storePathLabel.isBlank()) {
                    y = drawLine(cs, font, BODY_SIZE, MARGIN, y, "データ正本: " + storePathLabel.strip());
                }
                y -= LINE_HEIGHT * 0.5f;
                y = drawLine(cs, font, BODY_SIZE, MARGIN, y, "名前");
                drawLine(cs, font, BODY_SIZE, 180f, y + LINE_HEIGHT, "状態");
                drawLine(cs, font, BODY_SIZE, 280f, y + LINE_HEIGHT, "PIN（管理者閲覧）");
                y -= LINE_HEIGHT * 0.25f;
                cs.moveTo(MARGIN, y);
                cs.lineTo(page.getMediaBox().getWidth() - MARGIN, y);
                cs.stroke();
                y -= LINE_HEIGHT;
                for (Row row : rows) {
                    if (y < MARGIN + LINE_HEIGHT) {
                        throw new IOException("1 ページに収まりません。ユーザー数が多すぎます。");
                    }
                    y = drawLine(cs, font, BODY_SIZE, MARGIN, y, safe(row.name()));
                    drawLine(cs, font, BODY_SIZE, 180f, y + LINE_HEIGHT, safe(row.pinStatus()));
                    drawLine(cs, font, BODY_SIZE, 280f, y + LINE_HEIGHT, safe(row.adminPin()));
                }
                y -= LINE_HEIGHT;
                drawLine(
                        cs,
                        font,
                        BODY_SIZE,
                        MARGIN,
                        y,
                        "※ 管理者専用資料。PIN を含むため第三者への共有・放置に注意してください。");
            }
            document.save(path.toFile());
        }
    }

    private static float drawLine(
            PDPageContentStream cs, PDFont font, float size, float x, float y, String text)
            throws IOException {
        cs.beginText();
        cs.setFont(font, size);
        cs.newLineAtOffset(x, y);
        cs.showText(text != null ? text : "");
        cs.endText();
        return y - LINE_HEIGHT;
    }

    private static String safe(String raw) {
        return raw != null ? raw : "";
    }
}
