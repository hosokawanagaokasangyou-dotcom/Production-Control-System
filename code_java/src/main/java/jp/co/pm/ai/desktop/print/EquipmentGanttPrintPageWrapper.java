package jp.co.pm.ai.desktop.print;

import javafx.geometry.Bounds;
import javafx.scene.Node;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.SnapshotParameters;
import javafx.scene.canvas.Canvas;
import javafx.scene.canvas.GraphicsContext;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.image.WritableImage;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.transform.Scale;
import javafx.print.PageLayout;

import jp.co.pm.ai.desktop.ui.EquipmentGraphicGanttPane;

/**
 * 設備ガントを印刷用に「1 物理ページ」へ収める。
 *
 * <p>ガント本体（ScrollPane・多数 Canvas）は {@link javafx.print.PrinterJob#printPage} に渡さず、
 * レイアウト確定後に {@code snapshot} した 1 枚の画像だけを用紙サイズの {@link Canvas} に貼る。
 * 巨大ラスタや不可視 {@link javafx.stage.Stage} の乱立は JVM／PDF ドライバのクラッシュ・破損の原因になるため使わない。
 */
public final class EquipmentGanttPrintPageWrapper {

    /** スナップショット 1 辺の上限（Prism／Microsoft Print to PDF の安定域） */
    private static final double SNAPSHOT_MAX_EDGE = 4096;

    /** 総ピクセル数の上限（幅×高さ、OOM／ネイティブクラッシュ防止） */
    private static final long SNAPSHOT_MAX_PIXELS = 12_000_000L;

    private EquipmentGanttPrintPageWrapper() {}

    /**
     * ガント 1 日分を用紙の可印刷領域に収めた {@link Parent} を返す。
     *
     * @param gantt {@link EquipmentGraphicGanttPane#build} の戻り値
     * @param layout 用紙・向きが確定した {@link PageLayout}
     */
    public static Parent fitGanttToSinglePrintablePage(BorderPane gantt, PageLayout layout) {
        if (gantt == null || layout == null) {
            return new StackPane();
        }
        double paperW = layout.getPrintableWidth();
        double paperH = layout.getPrintableHeight();
        if (!Double.isFinite(paperW) || !Double.isFinite(paperH) || paperW < 2 || paperH < 2) {
            return new StackPane();
        }

        EquipmentGraphicGanttPane.EquipmentGanttViewHandles handles = viewHandles(gantt);
        prepareGanttForPrinting(gantt, handles);
        runFullTimelinePaint(handles);

        double cw = contentWidth(handles, gantt);
        double ch = contentHeight(handles, gantt);
        Node snapTarget = snapshotTarget(handles, gantt);

        pulseLayoutForPrint(gantt, handles, snapTarget, cw, ch);
        runFullTimelinePaint(handles);

        WritableImage img = captureSnapshot(snapTarget, cw, ch, paperW, paperH);
        if (snapshotLooksBlank(img) && snapTarget != gantt) {
            pulseLayoutForPrint(gantt, handles, gantt, cw, ch);
            runFullTimelinePaint(handles);
            img = captureSnapshot(gantt, cw, ch, paperW, paperH);
        }

        Parent paper;
        if (img != null && !snapshotLooksBlank(img)) {
            paper =
                    paperCanvasWithCenteredImage(
                            img, paperW, paperH, img.getWidth(), img.getHeight());
        } else {
            paper = errorPaper(paperW, paperH, "ガント画像の生成に失敗しました");
        }
        attachPrintScene(paper, paperW, paperH);
        return paper;
    }

    private static WritableImage captureSnapshot(
            Node snapTarget, double cw, double ch, double paperW, double paperH) {
        if (snapTarget == null) {
            return null;
        }
        double fitScale = Math.min(paperW / cw, paperH / ch);
        if (!Double.isFinite(fitScale) || fitScale <= 0) {
            fitScale = 1.0;
        }
        double rasterScale =
                Math.min(
                        SNAPSHOT_MAX_EDGE / cw,
                        Math.min(SNAPSHOT_MAX_EDGE / ch, fitScale));
        rasterScale = clampRasterScaleByPixelBudget(cw, ch, rasterScale);
        rasterScale = Math.max(1.0, rasterScale);

        int imgW = (int) Math.ceil(cw * rasterScale);
        int imgH = (int) Math.ceil(ch * rasterScale);
        imgW = (int) Math.min(imgW, SNAPSHOT_MAX_EDGE);
        imgH = (int) Math.min(imgH, SNAPSHOT_MAX_EDGE);
        imgW = Math.max(1, imgW);
        imgH = Math.max(1, imgH);

        SnapshotParameters snapParams = new SnapshotParameters();
        snapParams.setFill(Color.WHITE);
        if (rasterScale > 1.0 + 1e-9) {
            snapParams.setTransform(new Scale(rasterScale, rasterScale, 0, 0));
        }
        return trySnapshot(snapTarget, snapParams, imgW, imgH);
    }

    private static double clampRasterScaleByPixelBudget(double cw, double ch, double rasterScale) {
        if (cw < 1 || ch < 1) {
            return 1.0;
        }
        double pixels = cw * ch * rasterScale * rasterScale;
        if (pixels <= SNAPSHOT_MAX_PIXELS) {
            return rasterScale;
        }
        double allowed = Math.sqrt(SNAPSHOT_MAX_PIXELS / (cw * ch));
        return Math.min(rasterScale, Math.max(1.0, allowed));
    }

    private static WritableImage trySnapshot(
            Node target, SnapshotParameters params, int imgW, int imgH) {
        try {
            return target.snapshot(params, new WritableImage(imgW, imgH));
        } catch (RuntimeException | OutOfMemoryError ex) {
            return null;
        }
    }

    private static boolean snapshotLooksBlank(WritableImage img) {
        if (img == null || img.getPixelReader() == null) {
            return true;
        }
        int w = (int) img.getWidth();
        int h = (int) img.getHeight();
        if (w < 2 || h < 2) {
            return true;
        }
        int[] xs = {w / 4, w / 2, 3 * w / 4};
        int[] ys = {h / 4, h / 2, 3 * h / 4};
        int nonWhite = 0;
        var pr = img.getPixelReader();
        for (int y : ys) {
            for (int x : xs) {
                int argb = pr.getArgb(x, y);
                int a = (argb >>> 24) & 0xff;
                int r = (argb >>> 16) & 0xff;
                int g = (argb >>> 8) & 0xff;
                int b = argb & 0xff;
                if (a > 12 && (r < 245 || g < 245 || b < 245)) {
                    nonWhite++;
                }
            }
        }
        return nonWhite == 0;
    }

    private static void attachPrintScene(Parent printRoot, double paperW, double paperH) {
        if (printRoot == null) {
            return;
        }
        if (printRoot.getScene() == null) {
            new Scene(printRoot, paperW, paperH, Color.WHITE);
        }
        printRoot.applyCss();
        printRoot.layout();
    }

    private static Parent errorPaper(double paperW, double paperH, String message) {
        Label lab = new Label(message);
        lab.setWrapText(true);
        StackPane paper = new StackPane(lab);
        paper.setPrefSize(paperW, paperH);
        paper.setMinSize(paperW, paperH);
        paper.setMaxSize(paperW, paperH);
        paper.setStyle("-fx-background-color: white;");
        return paper;
    }

    private static EquipmentGraphicGanttPane.EquipmentGanttViewHandles viewHandles(BorderPane gantt) {
        Object ud = gantt.getUserData();
        if (ud instanceof EquipmentGraphicGanttPane.EquipmentGanttViewHandles h) {
            return h;
        }
        return null;
    }

    private static Node snapshotTarget(
            EquipmentGraphicGanttPane.EquipmentGanttViewHandles handles, BorderPane gantt) {
        if (handles != null && handles.mainColumn() != null) {
            return handles.mainColumn();
        }
        return gantt;
    }

    private static double contentWidth(
            EquipmentGraphicGanttPane.EquipmentGanttViewHandles handles, BorderPane gantt) {
        if (handles != null && handles.printContentWidth() > 1.0) {
            return handles.printContentWidth();
        }
        Bounds lb = gantt.getLayoutBounds();
        return Math.max(1.0, lb.getWidth());
    }

    private static double contentHeight(
            EquipmentGraphicGanttPane.EquipmentGanttViewHandles handles, BorderPane gantt) {
        if (handles != null && handles.printContentHeight() > 1.0) {
            return handles.printContentHeight();
        }
        Bounds lb = gantt.getLayoutBounds();
        return Math.max(1.0, lb.getHeight());
    }

    private static void runFullTimelinePaint(
            EquipmentGraphicGanttPane.EquipmentGanttViewHandles handles) {
        if (handles == null) {
            return;
        }
        Runnable fullPaint = handles.paintFullContentForPrint();
        if (fullPaint != null) {
            fullPaint.run();
            return;
        }
        Runnable repaint = handles.scheduleViewportRepaint();
        if (repaint != null) {
            repaint.run();
        }
    }

    private static void pulseLayoutForPrint(
            BorderPane gantt,
            EquipmentGraphicGanttPane.EquipmentGanttViewHandles handles,
            Node snapTarget,
            double cw,
            double ch) {
        if (gantt.getScene() == null) {
            new Scene(gantt, cw, ch, Color.WHITE);
        }
        gantt.applyCss();
        gantt.layout();
        snapTarget.applyCss();
        if (snapTarget instanceof Parent p) {
            p.layout();
        }
        runFullTimelinePaint(handles);
        gantt.applyCss();
        gantt.layout();
    }

    private static Parent paperCanvasWithCenteredImage(
            WritableImage img, double paperW, double paperH, double iw, double ih) {
        Canvas canvas = new Canvas(paperW, paperH);
        GraphicsContext gc = canvas.getGraphicsContext2D();
        gc.setFill(Color.WHITE);
        gc.fillRect(0, 0, paperW, paperH);
        double s = Math.min(paperW / iw, paperH / ih);
        double dw = iw * s;
        double dh = ih * s;
        double dx = (paperW - dw) * 0.5;
        double dy = (paperH - dh) * 0.5;
        gc.drawImage(img, dx, dy, dw, dh);

        StackPane paper = new StackPane(canvas);
        paper.setPrefSize(paperW, paperH);
        paper.setMinSize(paperW, paperH);
        paper.setMaxSize(paperW, paperH);
        return paper;
    }

    private static void prepareGanttForPrinting(
            BorderPane gantt, EquipmentGraphicGanttPane.EquipmentGanttViewHandles handles) {
        Node bottom = gantt.getBottom();
        if (bottom != null) {
            bottom.setVisible(false);
            bottom.setManaged(false);
        }
        if (handles == null) {
            return;
        }

        double cw = Math.max(1.0, handles.printContentWidth());
        double ch = Math.max(1.0, handles.printContentHeight());
        double headerH = headerRowHeight(handles);
        double bodyH = Math.max(1.0, ch - headerH - mainColumnVerticalPadding(handles));

        VBox main = handles.mainColumn();
        if (main != null) {
            hidePrintWarningBanner(main);
            VBox.setVgrow(handles.bodySplit(), Priority.NEVER);
            main.setPrefSize(cw, ch);
            main.setMinSize(cw, ch);
            main.setMaxSize(cw, ch);
        }

        HBox head = handles.headRow();
        if (head != null) {
            HBox.setHgrow(handles.headerScroll(), Priority.NEVER);
            head.setPrefSize(cw, headerH);
            head.setMinSize(cw, headerH);
            head.setMaxSize(cw, headerH);
        }

        HBox body = handles.bodySplit();
        if (body != null) {
            HBox.setHgrow(handles.timelineScroll(), Priority.NEVER);
            body.setPrefSize(cw, bodyH);
            body.setMinSize(cw, bodyH);
            body.setMaxSize(cw, bodyH);
        }

        gantt.setPrefSize(cw, ch);
        gantt.setMinSize(cw, ch);
        gantt.setMaxSize(cw, ch);

        resetScrollPaneOrigin(handles.leftBodyScroll());
        resetScrollPaneOrigin(handles.timelineScroll());
        resetScrollPaneOrigin(handles.headerScroll());
        expandScrollPaneToFullContent(handles.leftBodyScroll());
        expandScrollPaneToFullContent(handles.timelineScroll());
        expandScrollPaneToFullContent(handles.headerScroll());

        ScrollPane headerScroll = handles.headerScroll();
        if (headerScroll != null) {
            headerScroll.setFitToHeight(false);
        }

        gantt.applyCss();
        gantt.layout();
    }

    private static void hidePrintWarningBanner(VBox main) {
        if (main.getChildren().isEmpty()) {
            return;
        }
        Node first = main.getChildren().get(0);
        if (first instanceof Label) {
            first.setVisible(false);
            first.setManaged(false);
        }
    }

    private static double headerRowHeight(EquipmentGraphicGanttPane.EquipmentGanttViewHandles handles) {
        HBox head = handles.headRow();
        if (head != null) {
            double ph = head.getPrefHeight();
            if (Double.isFinite(ph) && ph > 0.5) {
                return ph;
            }
        }
        return Math.max(1.0, handles.printContentHeight() * 0.12);
    }

    private static double mainColumnVerticalPadding(
            EquipmentGraphicGanttPane.EquipmentGanttViewHandles handles) {
        VBox main = handles.mainColumn();
        if (main == null) {
            return 0.0;
        }
        return main.getPadding().getTop() + main.getPadding().getBottom();
    }

    private static void resetScrollPaneOrigin(ScrollPane sp) {
        if (sp == null) {
            return;
        }
        sp.setHvalue(0);
        sp.setVvalue(0);
    }

    private static void expandScrollPaneToFullContent(ScrollPane sp) {
        if (sp == null) {
            return;
        }
        Node c = sp.getContent();
        if (c == null) {
            return;
        }
        double w = Math.max(1.0, c.prefWidth(-1));
        double h = Math.max(1.0, c.prefHeight(-1));
        sp.setPrefViewportWidth(w);
        sp.setPrefViewportHeight(h);
        sp.setMinViewportWidth(w);
        sp.setMinViewportHeight(h);
        sp.setMaxWidth(w);
        sp.setPrefWidth(w);
        sp.setMinWidth(w);
        sp.setHbarPolicy(ScrollPane.ScrollBarPolicy.NEVER);
        sp.setVbarPolicy(ScrollPane.ScrollBarPolicy.NEVER);
    }
}
