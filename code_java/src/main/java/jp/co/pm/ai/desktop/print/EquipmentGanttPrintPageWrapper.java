package jp.co.pm.ai.desktop.print;

import javafx.geometry.Bounds;
import javafx.geometry.Pos;
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
import javafx.scene.layout.Region;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.transform.Scale;
import javafx.print.PageLayout;

import jp.co.pm.ai.desktop.ui.EquipmentGraphicGanttPane;

/**
 * 設備ガントを印刷用に「1 物理ページ」へ収める。
 *
 * <p>JavaFX の {@link javafx.print.PrinterJob#printPage} は、{@link Scene} に載っていない
 * {@link javafx.scene.canvas.Canvas} が真っ白になることがある。主経路はレイアウト確定後の高解像度
 * {@code snapshot} を用紙に貼り付け、返却する {@link Parent} を必ず {@link Scene} に載せる。
 * 左列・バッジはラスタでも十分な解像度を確保し、build 側の {@link EquipmentGraphicGanttPane#PRINT_LAYOUT_SCALE}
 * で Canvas 帯のピクセル数を増やす。
 */
public final class EquipmentGanttPrintPageWrapper {

    /** スナップショット 1 辺の安全上限（GPU／Prism の上限に合わせる） */
    private static final double SNAPSHOT_MAX_EDGE = 8192;

    /** 用紙への収まり倍率に対し、snapshot をどれだけ上乗せするか（鮮明化） */
    private static final double SNAPSHOT_SHARPNESS_FACTOR = 2.0;

    private EquipmentGanttPrintPageWrapper() {}

    /**
     * ガント 1 日分を用紙の可印刷領域に収めた {@link Parent} を返す（ダイアログで選んだ向き・余白を尊重）。
     *
     * @param gantt {@link EquipmentGraphicGanttPane#build} の戻り値（{@code highQualityPrint=true} 推奨）
     * @param layout 用紙・向きが確定した {@link PageLayout}
     */
    public static Parent fitGanttToSinglePrintablePage(BorderPane gantt, PageLayout layout) {
        if (gantt == null || layout == null) {
            return new StackPane();
        }
        double paperW = layout.getPrintableWidth();
        double paperH = layout.getPrintableHeight();
        if (!Double.isFinite(paperW) || !Double.isFinite(paperH) || paperW < 2 || paperH < 2) {
            return gantt;
        }

        EquipmentGraphicGanttPane.EquipmentGanttViewHandles handles = viewHandles(gantt);
        prepareGanttForPrinting(gantt, handles);
        runFullTimelinePaint(handles);

        double cw = contentWidth(handles, gantt);
        double ch = contentHeight(handles, gantt);
        Node snapTarget = snapshotTarget(handles, gantt);

        pulseLayoutForPrint(gantt, handles, snapTarget, cw, ch);
        runFullTimelinePaint(handles);

        double fitScale = Math.min(paperW / cw, paperH / ch);
        if (!Double.isFinite(fitScale) || fitScale <= 0) {
            fitScale = 1.0;
        }
        double rasterScale =
                Math.min(
                        SNAPSHOT_MAX_EDGE / cw,
                        Math.min(SNAPSHOT_MAX_EDGE / ch, fitScale * SNAPSHOT_SHARPNESS_FACTOR));
        rasterScale = Math.max(1.0, rasterScale);

        int imgW = (int) Math.ceil(Math.min(cw * rasterScale, SNAPSHOT_MAX_EDGE));
        int imgH = (int) Math.ceil(Math.min(ch * rasterScale, SNAPSHOT_MAX_EDGE));
        imgW = Math.max(1, imgW);
        imgH = Math.max(1, imgH);

        SnapshotParameters snapParams = new SnapshotParameters();
        snapParams.setFill(Color.WHITE);
        if (rasterScale > 1.0 + 1e-9) {
            snapParams.setTransform(new Scale(rasterScale, rasterScale, 0, 0));
        }

        WritableImage img;
        try {
            img = snapTarget.snapshot(snapParams, new WritableImage(imgW, imgH));
        } catch (RuntimeException ex) {
            return attachPrintScene(
                    vectorFitCenteredOnPaper(gantt, handles, paperW, paperH, cw, ch), paperW, paperH);
        }
        if (img == null || img.getPixelReader() == null || img.getWidth() < 1 || img.getHeight() < 1) {
            return attachPrintScene(
                    vectorFitCenteredOnPaper(gantt, handles, paperW, paperH, cw, ch), paperW, paperH);
        }

        return attachPrintScene(
                paperCanvasWithCenteredImage(img, paperW, paperH, img.getWidth(), img.getHeight()),
                paperW,
                paperH);
    }

    /**
     * {@link javafx.print.PrinterJob#printPage} 前に印刷ルートを {@link Scene} に載せ、レイアウトと Canvas 再描画を行う。
     */
    private static Parent attachPrintScene(Parent printRoot, double paperW, double paperH) {
        if (printRoot == null) {
            return new StackPane();
        }
        if (printRoot.getScene() == null) {
            new Scene(printRoot, paperW, paperH, Color.WHITE);
        }
        printRoot.applyCss();
        if (printRoot instanceof Parent p) {
            p.layout();
        }
        return printRoot;
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
        if (snapTarget.getScene() == null) {
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
        snapTarget.applyCss();
        if (snapTarget instanceof Parent p) {
            p.layout();
        }
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

    /**
     * snapshot 失敗時のフォールバック。{@link BorderPane} を用紙内にスケールして {@link StackPane} に載せる。
     */
    private static Parent vectorFitCenteredOnPaper(
            BorderPane gantt,
            EquipmentGraphicGanttPane.EquipmentGanttViewHandles handles,
            double paperW,
            double paperH,
            double contentW,
            double contentH) {
        double scale = Math.min(paperW / contentW, paperH / contentH);
        if (!Double.isFinite(scale) || scale <= 0) {
            scale = 1.0;
        }

        gantt.setPrefSize(contentW, contentH);
        gantt.setMinSize(contentW, contentH);
        gantt.setMaxSize(contentW, contentH);
        gantt.setScaleX(scale);
        gantt.setScaleY(scale);

        Region bg = new Region();
        bg.setMinSize(paperW, paperH);
        bg.setPrefSize(paperW, paperH);
        bg.setMaxSize(paperW, paperH);
        bg.setStyle("-fx-background-color: white;");

        StackPane paper = new StackPane(bg, gantt);
        StackPane.setAlignment(gantt, Pos.CENTER);
        paper.setPrefSize(paperW, paperH);
        paper.setMinSize(paperW, paperH);
        paper.setMaxSize(paperW, paperH);

        runFullTimelinePaint(handles);
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
            head.applyCss();
            head.layout();
            Bounds b = head.getLayoutBounds();
            if (b != null && b.getHeight() > 0.5) {
                return b.getHeight();
            }
        }
        ScrollPane headerScroll = handles.headerScroll();
        if (headerScroll != null) {
            Node c = headerScroll.getContent();
            if (c != null) {
                double h = c.prefHeight(-1);
                if (Double.isFinite(h) && h > 0.5) {
                    return h;
                }
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
