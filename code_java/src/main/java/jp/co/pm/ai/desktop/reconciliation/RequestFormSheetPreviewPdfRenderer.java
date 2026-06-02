package jp.co.pm.ai.desktop.reconciliation;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.reconciliation.RequestFormSheetPreviewRenderer.PreviewData;
import jp.co.pm.ai.desktop.reconciliation.RequestFormSheetShapeOverlay.OverlayShape;

/**
 * 依頼書 Excel 範囲（{@link RequestFormSheetPreviewRenderer#PREVIEW_RANGE_SPEC}）を Apache PDFBox で PDF 化する。
 * JavaFX スナップショットは使わず、POI で読んだセル値・罫線・シェイプを PDF に直接描画する。
 */
final class RequestFormSheetPreviewPdfRenderer {

    private static final String RENDERER_SPEC_BASE = "pdfbox-v5-unformatted-date";

    private static final float PX_TO_PT = 72f / 96f;
    private static volatile float cjkMetricsScale =
            AppPaths.DEFAULT_PM_AI_REQUEST_FORM_PREVIEW_PDF_CJK_SCALE;

    /** 描画方式の識別子（係数変更時もプレビューキャッシュを無効化する）。 */
    static String rendererSpec() {
        int pct = Math.round(cjkMetricsScale() * 100f);
        return RENDERER_SPEC_BASE + "-s" + pct;
    }

    static void applyCjkMetricsScaleFromUi(java.util.Map<String, String> ui) {
        cjkMetricsScale = AppPaths.resolveRequestFormPreviewPdfCjkScale(ui);
    }

    static float cjkMetricsScale() {
        return cjkMetricsScale;
    }
    private static final float PREVIEW_RENDER_DPI = 144f;
    private static final Pattern BORDER_WIDTH_PATTERN =
            Pattern.compile(
                    "-fx-border-width:\\s*([\\d.]+)\\s+([\\d.]+)\\s+([\\d.]+)\\s+([\\d.]+)\\s*;",
                    Pattern.CASE_INSENSITIVE);
    private static final Pattern BORDER_COLOR_PATTERN =
            Pattern.compile(
                    "-fx-border-color:\\s*(#[0-9A-Fa-f]{6})\\s+(#[0-9A-Fa-f]{6})\\s+(#[0-9A-Fa-f]{6})\\s+(#[0-9A-Fa-f]{6})\\s*;",
                    Pattern.CASE_INSENSITIVE);

    private RequestFormSheetPreviewPdfRenderer() {}

    static void generatePreviewPdf(File excelFile, String sheetName, File outputFile) throws Exception {
        PreviewData data = RequestFormSheetPreviewRenderer.loadPreviewData(excelFile, sheetName);
        writePdf(data, outputFile);
    }

    static BufferedImage renderFirstPageImage(File pdfFile, float dpi) throws IOException {
        try (PDDocument document = Loader.loadPDF(pdfFile)) {
            if (document.getNumberOfPages() <= 0) {
                throw new IOException("PDF にページがありません");
            }
            PDFRenderer renderer = new PDFRenderer(document);
            return renderer.renderImageWithDPI(0, dpi);
        }
    }

    static BufferedImage renderFirstPageImage(File pdfFile) throws IOException {
        return renderFirstPageImage(pdfFile, PREVIEW_RENDER_DPI);
    }

    private static void writePdf(PreviewData data, File outputFile) throws IOException {
        File parent = outputFile.getParentFile();
        if (parent != null) {
            Files.createDirectories(parent.toPath());
        }
        Path temp =
                Files.createTempFile(
                        parent != null ? parent.toPath() : Path.of("."),
                        "request-form-preview-",
                        ".pdf.tmp");
        try (PDDocument document = new PDDocument()) {
            RequestFormPreviewPdfFonts.FontPair fonts = RequestFormPreviewPdfFonts.load(document);
            float pageWidthPt = totalPx(data.colWidthsPx()) * PX_TO_PT;
            float pageHeightPt = totalPx(data.rowHeightsPx()) * PX_TO_PT;
            pageWidthPt = Math.max(72f, pageWidthPt);
            pageHeightPt = Math.max(72f, pageHeightPt);

            PDPage page = new PDPage(new PDRectangle(pageWidthPt, pageHeightPt));
            document.addPage(page);

            float[] colXPt = cumulativePxToPt(data.colWidthsPx());
            float[] rowYPt = cumulativePxToPt(data.rowHeightsPx());

            try (PDPageContentStream stream =
                    new PDPageContentStream(document, page, PDPageContentStream.AppendMode.OVERWRITE, true, true)) {
                stream.setStrokingColor(0f, 0f, 0f);
                stream.setNonStrokingColor(1f, 1f, 1f);
                stream.addRect(0f, 0f, pageWidthPt, pageHeightPt);
                stream.fill();

                drawCells(stream, document, fonts, data, colXPt, rowYPt, pageHeightPt);
                drawOverlayShapes(stream, document, fonts, data, pageHeightPt);
            }

            document.save(temp.toFile());
        }
        Files.move(temp, outputFile.toPath(), StandardCopyOption.REPLACE_EXISTING, StandardCopyOption.ATOMIC_MOVE);
    }

    private static void drawCells(
            PDPageContentStream stream,
            PDDocument document,
            RequestFormPreviewPdfFonts.FontPair fonts,
            PreviewData data,
            float[] colXPt,
            float[] rowYPt,
            float pageHeightPt)
            throws IOException {
        for (int r = 0; r < data.rowCount(); r++) {
            for (int c = 0; c < data.colCount(); c++) {
                if (data.skip()[r][c]) {
                    continue;
                }
                float xPt = colXPt[c];
                float yTopPt = rowYPt[r];
                float widthPt = spanWidthPt(data, r, c, colXPt);
                float heightPt = spanHeightPt(data, r, c, rowYPt);
                float yBottomPt = pageHeightPt - yTopPt - heightPt;

                RequestFormPreviewCellStyle style = data.styles()[r][c];
                if (style == null) {
                    style = RequestFormPreviewCellStyle.defaults();
                }
                fillBackground(stream, style, xPt, yBottomPt, widthPt, heightPt);
                drawBorder(stream, style.borderCss(), xPt, yBottomPt, widthPt, heightPt);

                List<RequestFormPreviewTextRun> runs = data.richRuns()[r][c];
                String text = data.texts()[r][c];
                if (runs != null && runs.size() > 1) {
                    drawRichText(
                            stream,
                            fonts,
                            runs,
                            data.hAligns()[r][c],
                            style.verticalAlignment(),
                            xPt,
                            yBottomPt,
                            widthPt,
                            heightPt);
                } else if (text != null && !text.isBlank()) {
                    drawPlainText(
                            stream,
                            fonts,
                            text,
                            style,
                            data.hAligns()[r][c],
                            xPt,
                            yBottomPt,
                            widthPt,
                            heightPt);
                }
            }
        }
    }

    private static void drawOverlayShapes(
            PDPageContentStream stream,
            PDDocument document,
            RequestFormPreviewPdfFonts.FontPair fonts,
            PreviewData data,
            float pageHeightPt)
            throws IOException {
        if (data.overlayShapes() == null) {
            return;
        }
        for (OverlayShape shape : data.overlayShapes()) {
            float xPt = (float) shape.x() * PX_TO_PT;
            float yTopPt = (float) shape.y() * PX_TO_PT;
            float widthPt = (float) shape.width() * PX_TO_PT;
            float heightPt = (float) shape.height() * PX_TO_PT;
            float yBottomPt = pageHeightPt - yTopPt - heightPt;

            if (shape.pictureBytes() != null && shape.pictureBytes().length > 0) {
                PDImageXObject image =
                        PDImageXObject.createFromByteArray(document, shape.pictureBytes(), "shape-picture");
                stream.drawImage(image, xPt, yBottomPt, widthPt, heightPt);
                continue;
            }

            if (!shape.noFill() && shape.fillHex() != null) {
                float[] rgb = hexToRgb(shape.fillHex());
                stream.setNonStrokingColor(rgb[0], rgb[1], rgb[2]);
                stream.addRect(xPt, yBottomPt, widthPt, heightPt);
                stream.fill();
            }
            if (!shape.noLine()) {
                float[] rgb = hexToRgb(shape.lineHex() != null ? shape.lineHex() : "#000000");
                stream.setStrokingColor(rgb[0], rgb[1], rgb[2]);
                stream.setLineWidth((float) Math.max(0.5, shape.lineWidthPx() * PX_TO_PT));
                stream.addRect(xPt, yBottomPt, widthPt, heightPt);
                stream.stroke();
            }
            if (shape.textRuns() != null && !shape.textRuns().isEmpty()) {
                drawRichText(
                        stream,
                        fonts,
                        shape.textRuns(),
                        HorizontalAlignment.CENTER,
                        VerticalAlignment.CENTER,
                        xPt,
                        yBottomPt,
                        widthPt,
                        heightPt);
            }
        }
    }

    private static void fillBackground(
            PDPageContentStream stream,
            RequestFormPreviewCellStyle style,
            float x,
            float y,
            float width,
            float height)
            throws IOException {
        String bg = style.background();
        if (bg == null || bg.isBlank() || "#FFFFFF".equalsIgnoreCase(bg)) {
            return;
        }
        float[] rgb = hexToRgb(bg);
        stream.setNonStrokingColor(rgb[0], rgb[1], rgb[2]);
        stream.addRect(x, y, width, height);
        stream.fill();
    }

    private static void drawBorder(
            PDPageContentStream stream, String borderCss, float x, float y, float width, float height)
            throws IOException {
        BorderSpec border = parseBorder(borderCss);
        if (border == null) {
            return;
        }
        if (border.top() > 0) {
            drawLine(stream, border.topColor(), x, y + height, x + width, y + height, border.top());
        }
        if (border.bottom() > 0) {
            drawLine(stream, border.bottomColor(), x, y, x + width, y, border.bottom());
        }
        if (border.left() > 0) {
            drawLine(stream, border.leftColor(), x, y, x, y + height, border.left());
        }
        if (border.right() > 0) {
            drawLine(stream, border.rightColor(), x + width, y, x + width, y + height, border.right());
        }
    }

    private static void drawLine(
            PDPageContentStream stream,
            String colorHex,
            float x1,
            float y1,
            float x2,
            float y2,
            double widthPx)
            throws IOException {
        float[] rgb = hexToRgb(colorHex != null ? colorHex : "#000000");
        stream.setStrokingColor(rgb[0], rgb[1], rgb[2]);
        stream.setLineWidth((float) Math.max(0.35, widthPx * PX_TO_PT));
        stream.moveTo(x1, y1);
        stream.lineTo(x2, y2);
        stream.stroke();
    }

    private static void drawPlainText(
            PDPageContentStream stream,
            RequestFormPreviewPdfFonts.FontPair fonts,
            String text,
            RequestFormPreviewCellStyle style,
            HorizontalAlignment hAlign,
            float x,
            float y,
            float width,
            float height)
            throws IOException {
        PDFont font = style.bold() ? fonts.bold() : fonts.regular();
        float padX = 3f * PX_TO_PT;
        float padY = Math.max(1f, Math.min(4f, (float) style.fontSizePx() * 0.08f)) * PX_TO_PT;
        float maxWidth = Math.max(1f, width - padX * 2f);
        float maxHeight = Math.max(1f, height - padY * 2f);
        float fontSizePt = fitFontSizeToBox(font, text, excelFontSizePt(style.fontSizePx()), maxWidth, maxHeight, style.wrapText());
        List<String> lines =
                style.wrapText()
                        ? wrapLines(text.replace("\r", ""), font, fontSizePt, maxWidth)
                        : List.of(text.replace("\r", "").replace('\n', ' '));
        float lineHeight = fontSizePt * 1.15f;
        float blockHeight = lineHeight * lines.size();
        float startY =
                switch (style.verticalAlignment()) {
                    case TOP, JUSTIFY -> y + height - padY - fontSizePt;
                    case BOTTOM, DISTRIBUTED -> y + padY + blockHeight - fontSizePt;
                    default -> y + (height - blockHeight) / 2f + blockHeight - fontSizePt;
                };
        float[] rgb = hexToRgb(style.foreground());
        stream.setNonStrokingColor(rgb[0], rgb[1], rgb[2]);
        float cursorY = startY;
        for (String line : lines) {
            float textWidth = stringWidth(font, line, fontSizePt);
            float textX =
                    switch (hAlign != null ? hAlign : HorizontalAlignment.LEFT) {
                        case CENTER, CENTER_SELECTION, FILL, JUSTIFY, DISTRIBUTED ->
                                x + (width - textWidth) / 2f;
                        case RIGHT -> x + width - padX - textWidth;
                        default -> x + padX;
                    };
            writeTextLine(stream, font, fontSizePt, textX, cursorY, line);
            drawTextStrike(stream, style, textX, cursorY, textWidth, fontSizePt);
            cursorY -= lineHeight;
            if (cursorY < y + padY) {
                break;
            }
        }
    }

    private static void drawRichText(
            PDPageContentStream stream,
            RequestFormPreviewPdfFonts.FontPair fonts,
            List<RequestFormPreviewTextRun> runs,
            HorizontalAlignment hAlign,
            VerticalAlignment vAlign,
            float x,
            float y,
            float width,
            float height)
            throws IOException {
        if (runs == null || runs.isEmpty()) {
            return;
        }
        float padX = 3f * PX_TO_PT;
        float padY = 2f * PX_TO_PT;
        float maxWidth = Math.max(1f, width - padX * 2f);
        float maxHeight = Math.max(1f, height - padY * 2f);
        RequestFormPreviewCellStyle base = runs.get(0).style();
        float fontSizePt =
                excelFontSizePt(base != null ? base.fontSizePx() : 11.0 * 96.0 / 72.0);
        float textWidth = 0f;
        float maxRunPt = fontSizePt;
        for (RequestFormPreviewTextRun run : runs) {
            PDFont runFont =
                    run.style() != null && run.style().bold() ? fonts.bold() : fonts.regular();
            float runSize =
                    excelFontSizePt(
                            run.style() != null ? run.style().fontSizePx() : 11.0 * 96.0 / 72.0);
            maxRunPt = Math.max(maxRunPt, runSize);
            textWidth += stringWidth(runFont, run.text() != null ? run.text() : "", runSize);
        }
        float widthScale = textWidth > maxWidth ? maxWidth / textWidth : 1f;
        float heightScale = maxRunPt > maxHeight ? maxHeight / maxRunPt : 1f;
        float runScale = Math.min(1f, Math.min(widthScale, heightScale));
        float scaledTextWidth = textWidth * runScale;
        fontSizePt *= runScale;
        float textX =
                switch (hAlign != null ? hAlign : HorizontalAlignment.LEFT) {
                    case CENTER, CENTER_SELECTION, FILL, JUSTIFY, DISTRIBUTED ->
                            x + (width - scaledTextWidth) / 2f;
                    case RIGHT -> x + width - padX - scaledTextWidth;
                    default -> x + padX;
                };
        float textY =
                switch (vAlign != null ? vAlign : VerticalAlignment.CENTER) {
                    case TOP, JUSTIFY -> y + height - padY - fontSizePt;
                    case BOTTOM, DISTRIBUTED -> y + padY;
                    default -> y + (height - fontSizePt) / 2f;
                };
        float cursorX = textX;
        for (RequestFormPreviewTextRun run : runs) {
            RequestFormPreviewCellStyle style = run.style();
            PDFont runFont = style != null && style.bold() ? fonts.bold() : fonts.regular();
            float runSize =
                    excelFontSizePt(
                                    style != null ? style.fontSizePx() : 11.0 * 96.0 / 72.0)
                            * runScale;
            float[] rgb = hexToRgb(style != null ? style.foreground() : "#000000");
            stream.setNonStrokingColor(rgb[0], rgb[1], rgb[2]);
            String text = run.text() != null ? run.text() : "";
            float runWidth = stringWidth(runFont, text, runSize);
            writeTextLine(stream, runFont, runSize, cursorX, textY, text);
            drawTextStrike(stream, style, cursorX, textY, runWidth, runSize);
            cursorX += runWidth;
        }
    }

    /**
     * 取り消し線（単線・二重線）をテキスト上に描画する。
     * Excel フォントの strikeout を PDF プレビューでも再現するため、文字描画の直後に呼ぶ。
     * {@code baselineY} は {@link #writeTextLine} に渡したベースライン（PDF 座標で下が小）。
     */
    private static void drawTextStrike(
            PDPageContentStream stream,
            RequestFormPreviewCellStyle style,
            float x,
            float baselineY,
            float textWidth,
            float fontSizePt)
            throws IOException {
        if (style == null || textWidth <= 0f || !(style.strike() || style.doubleStrike())) {
            return;
        }
        float[] rgb = hexToRgb(style.foreground());
        stream.setStrokingColor(rgb[0], rgb[1], rgb[2]);
        stream.setLineWidth(Math.max(0.4f, fontSizePt * 0.06f));
        float center = baselineY + fontSizePt * 0.28f;
        if (style.doubleStrike()) {
            float gap = Math.max(0.6f, fontSizePt * 0.09f);
            drawHLine(stream, x, center + gap, x + textWidth);
            drawHLine(stream, x, center - gap, x + textWidth);
        } else {
            drawHLine(stream, x, center, x + textWidth);
        }
    }

    private static void drawHLine(PDPageContentStream stream, float x1, float y, float x2)
            throws IOException {
        stream.moveTo(x1, y);
        stream.lineTo(x2, y);
        stream.stroke();
    }

    private static void writeTextLine(
            PDPageContentStream stream, PDFont font, float fontSize, float x, float y, String text)
            throws IOException {
        if (text == null || text.isEmpty()) {
            return;
        }
        stream.beginText();
        stream.setFont(font, fontSize);
        stream.newLineAtOffset(x, y);
        stream.showText(text);
        stream.endText();
    }

    private static float excelFontSizePt(double fontSizePx) {
        return (float) Math.max(6.0, fontSizePx * PX_TO_PT * cjkMetricsScale());
    }

    private static float fitFontSizeToBox(
            PDFont font,
            String text,
            float fontSizePt,
            float maxWidth,
            float maxHeight,
            boolean wrapText)
            throws IOException {
        if (text == null || text.isBlank()) {
            return Math.min(fontSizePt, Math.max(6f, maxHeight * 0.92f));
        }
        String normalized = text.replace("\r", "");
        if (!wrapText) {
            String line = normalized.replace('\n', ' ');
            fontSizePt = fitFontSizeToWidth(font, line, fontSizePt, maxWidth);
            return Math.min(fontSizePt, Math.max(6f, maxHeight * 0.92f));
        }
        List<String> lines = wrapLines(normalized, font, fontSizePt, maxWidth);
        for (String line : lines) {
            fontSizePt = fitFontSizeToWidth(font, line, fontSizePt, maxWidth);
        }
        float lineHeight = fontSizePt * 1.15f;
        float blockHeight = lineHeight * Math.max(1, lines.size());
        if (blockHeight > maxHeight && blockHeight > 0f) {
            fontSizePt = Math.max(6f, fontSizePt * maxHeight / blockHeight);
        }
        return fontSizePt;
    }

    private static float fitFontSizeToWidth(PDFont font, String text, float fontSizePt, float maxWidth)
            throws IOException {
        if (text == null || text.isEmpty() || maxWidth <= 0f) {
            return fontSizePt;
        }
        float width = stringWidth(font, text, fontSizePt);
        if (width <= maxWidth || width <= 0f) {
            return fontSizePt;
        }
        return Math.max(6f, fontSizePt * maxWidth / width);
    }

    private static float stringWidth(PDFont font, String text, float fontSize) throws IOException {
        if (text == null || text.isEmpty()) {
            return 0f;
        }
        return font.getStringWidth(text) / 1000f * fontSize;
    }

    private static List<String> wrapLines(String text, PDFont font, float fontSize, float maxWidth)
            throws IOException {
        List<String> lines = new ArrayList<>();
        if (text == null || text.isEmpty()) {
            return lines;
        }
        for (String paragraph : text.split("\n", -1)) {
            if (paragraph.isEmpty()) {
                lines.add("");
                continue;
            }
            StringBuilder current = new StringBuilder();
            for (int i = 0; i < paragraph.length(); i++) {
                char ch = paragraph.charAt(i);
                String candidate = current.toString() + ch;
                if (stringWidth(font, candidate, fontSize) <= maxWidth || current.isEmpty()) {
                    current.append(ch);
                } else {
                    lines.add(current.toString());
                    current.setLength(0);
                    current.append(ch);
                }
            }
            if (!current.isEmpty()) {
                lines.add(current.toString());
            }
        }
        return lines;
    }

    private record BorderSpec(
            double top,
            double right,
            double bottom,
            double left,
            String topColor,
            String rightColor,
            String bottomColor,
            String leftColor) {}

    private static BorderSpec parseBorder(String borderCss) {
        if (borderCss == null || borderCss.isBlank()) {
            return null;
        }
        Matcher widthMatcher = BORDER_WIDTH_PATTERN.matcher(borderCss);
        Matcher colorMatcher = BORDER_COLOR_PATTERN.matcher(borderCss);
        if (!widthMatcher.find() || !colorMatcher.find()) {
            return null;
        }
        return new BorderSpec(
                Double.parseDouble(widthMatcher.group(1)),
                Double.parseDouble(widthMatcher.group(2)),
                Double.parseDouble(widthMatcher.group(3)),
                Double.parseDouble(widthMatcher.group(4)),
                colorMatcher.group(1),
                colorMatcher.group(2),
                colorMatcher.group(3),
                colorMatcher.group(4));
    }

    private static float[] hexToRgb(String hex) {
        if (hex == null || hex.length() < 7 || !hex.startsWith("#")) {
            return new float[] {0f, 0f, 0f};
        }
        int r = Integer.parseInt(hex.substring(1, 3), 16);
        int g = Integer.parseInt(hex.substring(3, 5), 16);
        int b = Integer.parseInt(hex.substring(5, 7), 16);
        return new float[] {r / 255f, g / 255f, b / 255f};
    }

    private static float totalPx(double[] values) {
        float total = 0f;
        if (values != null) {
            for (double value : values) {
                total += (float) value;
            }
        }
        return total;
    }

    private static float[] cumulativePxToPt(double[] values) {
        float[] result = new float[values.length];
        float cursor = 0f;
        for (int i = 0; i < values.length; i++) {
            result[i] = cursor;
            cursor += (float) values[i] * PX_TO_PT;
        }
        return result;
    }

    private static float spanWidthPt(PreviewData data, int row, int col, float[] colXPt) {
        int span = Math.max(1, data.colSpans()[row][col]);
        int end = Math.min(data.colCount(), col + span);
        float start = colXPt[col];
        float endX = colXPt[end - 1] + (float) data.colWidthsPx()[end - 1] * PX_TO_PT;
        return Math.max(1f, endX - start);
    }

    private static float spanHeightPt(PreviewData data, int row, int col, float[] rowYPt) {
        int span = Math.max(1, data.rowSpans()[row][col]);
        int end = Math.min(data.rowCount(), row + span);
        float start = rowYPt[row];
        float endY = rowYPt[end - 1] + (float) data.rowHeightsPx()[end - 1] * PX_TO_PT;
        return Math.max(1f, endY - start);
    }
}
