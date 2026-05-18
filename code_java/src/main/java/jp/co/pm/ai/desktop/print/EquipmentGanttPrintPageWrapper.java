package jp.co.pm.ai.desktop.print;

import javafx.geometry.Bounds;
import javafx.geometry.Pos;
import javafx.scene.Group;
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
 * <p>画面上は {@link ScrollPane} のビューポート＋部分描画のため、印刷直前にビューポートを内容全体に広げ全再描画し、
 * {@code snapshot} でラスタ化する。{@link BorderPane#getLayoutBounds()} は {@link HBox} の {@code hgrow} により
 * 実コンテンツより大きくなりがちなため、{@link EquipmentGraphicGanttPane.EquipmentGanttViewHandles} が保持する
 * 印刷寸法でシーンを切り、{@link #snapshotTarget} は余白の入らない {@link VBox#mainColumn} に限定する。
 */
public final class EquipmentGanttPrintPageWrapper {

    /** スナップショット 1 辺の安全上限（GPU／Prism の上限に合わせる） */
    private static final double SNAPSHOT_MAX_EDGE = 8192;

    private EquipmentGanttPrintPageWrapper() {}

    /**
     * ガント 1 日分を A3 横の可印刷領域に収めた {@link Parent} を返す。
     *
     * @param gantt {@link EquipmentGraphicGanttPane#build} の戻り値
     * @param layout 用紙・向きが確定した {@link PageLayout}（横向き想定）
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
        if (paperH > paperW + 0.5) {
            double t = paperW;
            paperW = paperH;
            paperH = t;
        }

        EquipmentGraphicGanttPane.EquipmentGanttViewHandles handles = viewHandles(gantt);
        prepareGanttForPrinting(gantt, handles);
        runFullTimelinePaint(handles);

        double bw = contentWidthForSnapshot(handles, gantt);
        double bh = contentHeightForSnapshot(handles, gantt);
        Node snapTarget = snapshotTarget(handles, gantt);

        double preScale = 1.0;
        if (bw > SNAPSHOT_MAX_EDGE || bh > SNAPSHOT_MAX_EDGE) {
            preScale = Math.min(SNAPSHOT_MAX_EDGE / bw, SNAPSHOT_MAX_EDGE / bh);
        }
        int imgW = (int) Math.ceil(bw * preScale);
        int imgH = (int) Math.ceil(bh * preScale);
        imgW = Math.max(1, Math.min(imgW, (int) SNAPSHOT_MAX_EDGE));
        imgH = Math.max(1, Math.min(imgH, (int) SNAPSHOT_MAX_EDGE));

        Scene shotScene = new Scene(gantt, bw, bh, Color.WHITE);
        java.util.Objects.requireNonNull(shotScene, "scene");
        gantt.applyCss();
        gantt.layout();
        runFullTimelinePaint(handles);
        gantt.applyCss();
        gantt.layout();
        snapTarget.applyCss();
        if (snapTarget instanceof Parent p) {
            p.layout();
        }

        SnapshotParameters snapParams = new SnapshotParameters();
        snapParams.setFill(Color.WHITE);
        if (preScale < 1.0 - 1e-9) {
            snapParams.setTransform(new Scale(preScale, preScale, 0, 0));
        }

        WritableImage img;
        try {
            img = snapTarget.snapshot(snapParams, new WritableImage(imgW, imgH));
        } catch (RuntimeException ex) {
            return vectorFallbackFit(gantt, paperW, paperH, bw, bh);
        }
        if (img == null || img.getPixelReader() == null) {
            return vectorFallbackFit(gantt, paperW, paperH, bw, bh);
        }

        double iw = img.getWidth();
        double ih = img.getHeight();
        if (iw < 1 || ih < 1) {
            return vectorFallbackFit(gantt, paperW, paperH, bw, bh);
        }

        return paperCanvasWithCenteredImage(img, paperW, paperH, iw, ih);
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

    private static double contentWidthForSnapshot(
            EquipmentGraphicGanttPane.EquipmentGanttViewHandles handles, BorderPane gantt) {
        if (handles != null && handles.printContentWidth() > 1.0) {
            return handles.printContentWidth();
        }
        Bounds lb = gantt.getLayoutBounds();
        return Math.max(1.0, lb.getWidth());
    }

    private static double contentHeightForSnapshot(
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

    /** snapshot 失敗時のみ。 */
    private static Parent vectorFallbackFit(
            BorderPane gantt, double paperW, double paperH, double bw, double bh) {
        double scale = Math.min(paperW / bw, paperH / bh);
        Group holder = new Group(gantt);
        gantt.setLayoutX(0);
        gantt.setLayoutY(0);
        if (scale < 1.0 - 1e-9) {
            holder.getTransforms().add(new Scale(scale, scale, 0, 0));
        }
        double dw = bw * Math.min(scale, 1.0);
        double dh = bh * Math.min(scale, 1.0);
        double dx = (paperW - dw) * 0.5;
        double dy = (paperH - dh) * 0.5;
        holder.setLayoutX(dx);
        holder.setLayoutY(dy);
        StackPane paper = new StackPane(holder);
        paper.setAlignment(Pos.CENTER);
        paper.setPrefSize(paperW, paperH);
        paper.setMinSize(paperW, paperH);
        paper.setMaxSize(paperW, paperH);
        return paper;
    }
}
