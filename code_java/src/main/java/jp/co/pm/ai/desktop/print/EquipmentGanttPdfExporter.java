package jp.co.pm.ai.desktop.print;

import java.awt.image.BufferedImage;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import javafx.embed.swing.SwingFXUtils;
import javafx.geometry.Rectangle2D;
import javafx.print.PageLayout;
import javafx.print.PageOrientation;
import javafx.print.Paper;
import javafx.print.Printer;
import javafx.scene.Parent;
import javafx.scene.SnapshotParameters;
import javafx.scene.image.WritableImage;
import javafx.scene.paint.Color;
import javafx.scene.transform.Transform;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.graphics.image.LosslessFactory;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;

import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * 設備ガント印刷専用レイアウト（{@link EquipmentGanttPrintCompositor}）を Apache PDFBox で PDF 化する。
 * Microsoft Print to PDF 等の仮想プリンターは使わない。
 */
public final class EquipmentGanttPdfExporter {

    /** 72pt 基準のラスター倍率（A3 横でも文字が潰れにくい程度）。 */
    private static final double SNAPSHOT_SCALE = 2.0;

    private EquipmentGanttPdfExporter() {}

    /** {@link AppPaths#equipmentGanttPdfPath(Map)} と同一の出力先。 */
    public static Path resolveOutputPath(Map<String, String> ui) {
        return AppPaths.equipmentGanttPdfPath(ui);
    }

    /** 設備ガント印刷と同じ A3 横向き・余白最小。 */
    public static PageLayout defaultPageLayout() throws IOException {
        Printer printer = Printer.getDefaultPrinter();
        if (printer == null) {
            throw new IOException("既定プリンターが無いため A3 横のページサイズを決定できません。");
        }
        return printer.createPageLayout(
                Paper.A3, PageOrientation.LANDSCAPE, Printer.MarginType.HARDWARE_MINIMUM);
    }

    /**
     * 複数ページを 1 ファイルへ原子的に書き出す。
     *
     * @param outputPath 上書き先（親フォルダは存在しなくてよい）
     * @param layout 各ページの可印刷領域（pt）
     * @param pageRoots {@link EquipmentGanttPrintCompositor#composePage} で組み立てたルート
     */
    public static void export(Path outputPath, PageLayout layout, List<Parent> pageRoots)
            throws IOException {
        if (outputPath == null) {
            throw new IOException("output path is null");
        }
        if (layout == null) {
            throw new IOException("page layout is null");
        }
        if (pageRoots == null || pageRoots.isEmpty()) {
            throw new IOException("印刷ページがありません");
        }
        Path parent = outputPath.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        Path temp =
                Files.createTempFile(
                        parent != null ? parent : Path.of("."),
                        "equipment-gantt-",
                        ".pdf.tmp");
        try {
            writePdf(temp, layout, pageRoots);
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

    private static void writePdf(Path path, PageLayout layout, List<Parent> pageRoots)
            throws IOException {
        double widthPt = layout.getPrintableWidth();
        double heightPt = layout.getPrintableHeight();
        if (!Double.isFinite(widthPt)
                || !Double.isFinite(heightPt)
                || widthPt < 2
                || heightPt < 2) {
            throw new IOException("不正な可印刷領域: " + widthPt + " x " + heightPt);
        }

        List<BufferedImage> rasters = new ArrayList<>(pageRoots.size());
        for (Parent root : pageRoots) {
            rasters.add(rasterizePage(root, widthPt, heightPt));
        }

        try (PDDocument document = new PDDocument()) {
            PDRectangle mediaBox = new PDRectangle((float) widthPt, (float) heightPt);
            for (BufferedImage raster : rasters) {
                PDPage page = new PDPage(mediaBox);
                document.addPage(page);
                PDImageXObject image = LosslessFactory.createFromImage(document, raster);
                try (PDPageContentStream contentStream =
                        new PDPageContentStream(document, page)) {
                    contentStream.drawImage(
                            image, 0f, 0f, (float) widthPt, (float) heightPt);
                }
            }
            document.save(path.toFile());
        }
    }

    private static BufferedImage rasterizePage(Parent root, double widthPt, double heightPt) {
        root.applyCss();
        root.layout();
        SnapshotParameters params = new SnapshotParameters();
        params.setFill(Color.WHITE);
        params.setViewport(new Rectangle2D(0, 0, widthPt, heightPt));
        params.setTransform(Transform.scale(SNAPSHOT_SCALE, SNAPSHOT_SCALE));
        int pixelW = (int) Math.ceil(widthPt * SNAPSHOT_SCALE);
        int pixelH = (int) Math.ceil(heightPt * SNAPSHOT_SCALE);
        WritableImage image = new WritableImage(pixelW, pixelH);
        root.snapshot(params, image);
        BufferedImage buffered = SwingFXUtils.fromFXImage(image, null);
        if (buffered == null) {
            throw new IllegalStateException("JavaFX snapshot failed");
        }
        return buffered;
    }
}
