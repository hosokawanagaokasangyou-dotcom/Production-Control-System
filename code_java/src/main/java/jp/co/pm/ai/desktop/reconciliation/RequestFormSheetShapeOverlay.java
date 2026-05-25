package jp.co.pm.ai.desktop.reconciliation;

import java.io.ByteArrayInputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFShapeGroup;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTextParagraph;
import org.apache.poi.xssf.usermodel.XSSFTextRun;

import javafx.geometry.Pos;
import javafx.scene.Node;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.Pane;
import javafx.scene.layout.StackPane;
import javafx.scene.shape.Rectangle;
import javafx.scene.shape.StrokeType;

/** 依頼書シート上のオートシェイプを JavaFX ノードへ変換する。 */
final class RequestFormSheetShapeOverlay {

    record OverlayShape(
            double x,
            double y,
            double width,
            double height,
            String fillHex,
            String lineHex,
            double lineWidthPx,
            boolean noFill,
            boolean noLine,
            List<RequestFormPreviewTextRun> textRuns,
            byte[] pictureBytes) {}

    private RequestFormSheetShapeOverlay() {}

    static List<OverlayShape> loadShapes(
            XSSFSheet sheet,
            int firstRow,
            int lastRow,
            int firstCol,
            int lastCol,
            double[] colWidthsPx,
            double[] rowHeightsPx) {
        List<OverlayShape> shapes = new ArrayList<>();
        XSSFDrawing drawing = sheet.getDrawingPatriarch();
        if (drawing == null) {
            return shapes;
        }
        collectShapes(
                drawing.getShapes(),
                shapes,
                firstRow,
                lastRow,
                firstCol,
                lastCol,
                colWidthsPx,
                rowHeightsPx);
        return shapes;
    }

    static Pane buildPane(List<OverlayShape> shapes) {
        Pane pane = new Pane();
        if (shapes == null) {
            return pane;
        }
        for (OverlayShape shape : shapes) {
            Node node = toNode(shape);
            if (node != null) {
                pane.getChildren().add(node);
            }
        }
        return pane;
    }

    private static void collectShapes(
            List<XSSFShape> source,
            List<OverlayShape> out,
            int firstRow,
            int lastRow,
            int firstCol,
            int lastCol,
            double[] colWidthsPx,
            double[] rowHeightsPx) {
        if (source == null) {
            return;
        }
        for (XSSFShape shape : source) {
            if (shape instanceof XSSFShapeGroup group) {
                List<XSSFShape> children = new ArrayList<>();
                group.iterator().forEachRemaining(children::add);
                collectShapes(
                        children,
                        out,
                        firstRow,
                        lastRow,
                        firstCol,
                        lastCol,
                        colWidthsPx,
                        rowHeightsPx);
                continue;
            }
            if (!(shape.getAnchor() instanceof XSSFClientAnchor anchor) || !anchor.isSet()) {
                continue;
            }
            if (isHidden(shape)) {
                continue;
            }
            double[] bounds =
                    anchorBounds(
                            anchor, firstRow, lastRow, firstCol, lastCol, colWidthsPx, rowHeightsPx);
            if (bounds == null) {
                continue;
            }
            double x = bounds[0];
            double y = bounds[1];
            double w = Math.max(1.0, bounds[2] - bounds[0]);
            double h = Math.max(1.0, bounds[3] - bounds[1]);

            if (shape instanceof XSSFPicture picture) {
                var data = picture.getPictureData();
                if (data != null && data.getData() != null && data.getData().length > 0) {
                    out.add(
                            new OverlayShape(
                                    x, y, w, h, null, null, 0, true, true, List.of(), data.getData()));
                }
                continue;
            }
            if (shape instanceof XSSFSimpleShape simple) {
                List<RequestFormPreviewTextRun> runs;
                try {
                    runs = textRunsFromShape(simple);
                } catch (RuntimeException ex) {
                    runs = List.of();
                }
                boolean hasText = !runs.isEmpty();
                String fill = null;
                if (!hasText && !simple.isNoFill()) {
                    fill = "#FFFFFF";
                }
                String line = "#000000";
                double lineWidth = 1.0;
                out.add(
                        new OverlayShape(
                                x,
                                y,
                                w,
                                h,
                                fill,
                                line,
                                lineWidth,
                                fill == null,
                                hasText || simple.isNoFill(),
                                runs,
                                null));
            }
        }
    }

    private static boolean isHidden(XSSFShape shape) {
        try {
            if (shape instanceof XSSFSimpleShape simple) {
                var nv = simple.getCTShape().getNvSpPr();
                if (nv != null && nv.getCNvPr() != null && nv.getCNvPr().isSetHidden()) {
                    return nv.getCNvPr().getHidden();
                }
            }
            if (shape instanceof XSSFPicture picture) {
                var nv = picture.getCTPicture().getNvPicPr();
                if (nv != null && nv.getCNvPr() != null && nv.getCNvPr().isSetHidden()) {
                    return nv.getCNvPr().getHidden();
                }
            }
        } catch (Exception ignored) {
            return false;
        }
        return false;
    }

    private static List<RequestFormPreviewTextRun> textRunsFromShape(XSSFSimpleShape shape) {
        List<RequestFormPreviewTextRun> runs = new ArrayList<>();
        for (XSSFTextParagraph paragraph : shape.getTextParagraphs()) {
            for (XSSFTextRun run : paragraph.getTextRuns()) {
                String text = run.getText();
                if (text == null || text.isEmpty()) {
                    continue;
                }
                runs.add(
                        new RequestFormPreviewTextRun(
                                text, RequestFormPreviewStyleHelper.fromTextRun(run, null)));
            }
        }
        if (runs.isEmpty()) {
            String plain = shape.getText();
            if (plain != null && !plain.isBlank()) {
                runs.add(
                        new RequestFormPreviewTextRun(
                                plain, RequestFormPreviewCellStyle.defaults()));
            }
        }
        return runs;
    }

    private static double[] anchorBounds(
            XSSFClientAnchor anchor,
            int firstRow,
            int lastRow,
            int firstCol,
            int lastCol,
            double[] colWidthsPx,
            double[] rowHeightsPx) {
        double x1 = cellOffsetX(anchor.getCol1(), anchor.getDx1(), firstCol, colWidthsPx);
        double y1 = cellOffsetY(anchor.getRow1(), anchor.getDy1(), firstRow, rowHeightsPx);
        double x2 = cellOffsetX(anchor.getCol2(), anchor.getDx2(), firstCol, colWidthsPx);
        double y2 = cellOffsetY(anchor.getRow2(), anchor.getDy2(), firstRow, rowHeightsPx);
        if (x2 < x1) {
            double t = x1;
            x1 = x2;
            x2 = t;
        }
        if (y2 < y1) {
            double t = y1;
            y1 = y2;
            y2 = t;
        }
        double previewW = sum(colWidthsPx);
        double previewH = sum(rowHeightsPx);
        if (x2 <= 0 || y2 <= 0 || x1 >= previewW || y1 >= previewH) {
            return null;
        }
        return new double[] {Math.max(0, x1), Math.max(0, y1), Math.min(previewW, x2), Math.min(previewH, y2)};
    }

    private static double cellOffsetX(int col, int dxEmu, int firstCol, double[] colWidthsPx) {
        double x = 0.0;
        for (int c = firstCol; c < col && c - firstCol < colWidthsPx.length; c++) {
            x += colWidthsPx[c - firstCol];
        }
        if (col >= firstCol) {
            x += (double) dxEmu / Units.EMU_PER_PIXEL;
        }
        return x;
    }

    private static double cellOffsetY(int row, int dyEmu, int firstRow, double[] rowHeightsPx) {
        double y = 0.0;
        for (int r = firstRow; r < row && r - firstRow < rowHeightsPx.length; r++) {
            y += rowHeightsPx[r - firstRow];
        }
        if (row >= firstRow) {
            y += (double) dyEmu / Units.EMU_PER_PIXEL;
        }
        return y;
    }

    private static double sum(double[] values) {
        double total = 0.0;
        if (values != null) {
            for (double v : values) {
                total += v;
            }
        }
        return total;
    }

    private static Node toNode(OverlayShape shape) {
        if (shape.pictureBytes() != null && shape.pictureBytes().length > 0) {
            ImageView view = new ImageView(new Image(new ByteArrayInputStream(shape.pictureBytes())));
            view.setFitWidth(shape.width());
            view.setFitHeight(shape.height());
            view.setPreserveRatio(true);
            layoutNode(view, shape);
            return view;
        }
        StackPane stack = new StackPane();
        stack.setMinSize(shape.width(), shape.height());
        stack.setPrefSize(shape.width(), shape.height());
        stack.setMaxSize(shape.width(), shape.height());
        stack.setMouseTransparent(!shape.textRuns().isEmpty());
        if (!shape.textRuns().isEmpty()) {
            Node textNode =
                    RequestFormPreviewNodeFactory.buildTextRuns(
                            shape.textRuns(), shape.width(), shape.height());
            StackPane.setAlignment(textNode, Pos.CENTER);
            stack.getChildren().add(textNode);
        } else {
            if (!shape.noFill()) {
                Rectangle fill = new Rectangle(shape.width(), shape.height());
                fill.setFill(
                        RequestFormPreviewNodeFactory.color(
                                shape.fillHex(), javafx.scene.paint.Color.WHITE));
                fill.setStrokeType(StrokeType.INSIDE);
                if (!shape.noLine()) {
                    fill.setStroke(
                            RequestFormPreviewNodeFactory.color(
                                    shape.lineHex(), javafx.scene.paint.Color.BLACK));
                    fill.setStrokeWidth(shape.lineWidthPx());
                }
                stack.getChildren().add(fill);
            } else if (!shape.noLine()) {
                Rectangle border = new Rectangle(shape.width(), shape.height());
                border.setFill(javafx.scene.paint.Color.TRANSPARENT);
                border.setStroke(
                        RequestFormPreviewNodeFactory.color(
                                shape.lineHex(), javafx.scene.paint.Color.BLACK));
                border.setStrokeWidth(shape.lineWidthPx());
                stack.getChildren().add(border);
            }
        }
        layoutNode(stack, shape);
        return stack;
    }

    private static void layoutNode(Node node, OverlayShape shape) {
        node.setLayoutX(shape.x());
        node.setLayoutY(shape.y());
    }
}
