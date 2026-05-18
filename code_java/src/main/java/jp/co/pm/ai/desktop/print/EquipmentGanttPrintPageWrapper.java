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
import javafx.scene.control.ScrollPane;
import javafx.scene.image.WritableImage;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.StackPane;
import javafx.scene.paint.Color;
import javafx.scene.transform.Scale;
import javafx.print.PageLayout;

import jp.co.pm.ai.desktop.ui.EquipmentGraphicGanttPane;

/**
 * 設備ガントを印刷用に「1 物理ページ」へ収める。
 *
 * <p>画面上は {@link ScrollPane} のビューポート＋部分描画のため、印刷直前にビューポートを内容全体に広げ全再描画し、
 * {@code snapshot} でラスタ化する。シーンを過大にすると白余白だけの巨大画像になり縮小で豆粒になるため、
 * レイアウト確定後のガント寸法に合わせたシーンで撮影する。用紙上は {@link Canvas} に等比拡大して中央配置する。
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
        /*
         * A3 横向きでは可印刷幅＞高さのはず。一部ドライバで逆転する場合に備え、長辺を横として扱う。
         */
        if (paperH > paperW + 0.5) {
            double t = paperW;
            paperW = paperH;
            paperH = t;
        }

        Scene probe = new Scene(gantt, 1200, 800, Color.WHITE);
        java.util.Objects.requireNonNull(probe, "scene");
        prepareGanttForPrinting(gantt);
        gantt.applyCss();
        gantt.layout();
        Runnable fullPaint = extractFullPaint(gantt);
        if (fullPaint != null) {
            fullPaint.run();
        } else {
            Runnable repaint = extractViewportRepaint(gantt);
            if (repaint != null) {
                repaint.run();
            }
        }
        gantt.applyCss();
        gantt.layout();

        Bounds lb = gantt.getLayoutBounds();
        double bw = Math.max(1.0, lb.getWidth());
        double bh = Math.max(1.0, lb.getHeight());

        probe.setRoot(new StackPane());

        double preScale = 1.0;
        if (bw > SNAPSHOT_MAX_EDGE || bh > SNAPSHOT_MAX_EDGE) {
            preScale = Math.min(SNAPSHOT_MAX_EDGE / bw, SNAPSHOT_MAX_EDGE / bh);
        }
        int imgW = (int) Math.ceil(bw * preScale);
        int imgH = (int) Math.ceil(bh * preScale);
        imgW = Math.max(1, Math.min(imgW, (int) SNAPSHOT_MAX_EDGE));
        imgH = Math.max(1, Math.min(imgH, (int) SNAPSHOT_MAX_EDGE));

        Scene shotScene = new Scene(gantt, bw, bh, Color.WHITE);
        gantt.applyCss();
        gantt.layout();
        if (fullPaint != null) {
            fullPaint.run();
        } else {
            Runnable repaint = extractViewportRepaint(gantt);
            if (repaint != null) {
                repaint.run();
            }
        }
        gantt.applyCss();
        gantt.layout();

        SnapshotParameters snapParams = new SnapshotParameters();
        snapParams.setFill(Color.WHITE);
        if (preScale < 1.0 - 1e-9) {
            snapParams.setTransform(new Scale(preScale, preScale, 0, 0));
        }

        WritableImage img;
        try {
            img = gantt.snapshot(snapParams, new WritableImage(imgW, imgH));
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

        Canvas canvas = new Canvas(paperW, paperH);
        GraphicsContext gc = canvas.getGraphicsContext2D();
        gc.setFill(Color.WHITE);
        gc.fillRect(0, 0, paperW, paperH);
        double s = Math.min(paperW / iw, paperH / ih);
        double dw = iw * s;
        double dh = ih * s;
        gc.drawImage(img, 0, 0, dw, dh);

        StackPane paper = new StackPane(canvas);
        paper.setPrefSize(paperW, paperH);
        paper.setMinSize(paperW, paperH);
        paper.setMaxSize(paperW, paperH);
        return paper;
    }

    private static Runnable extractFullPaint(BorderPane gantt) {
        Object ud = gantt.getUserData();
        if (ud instanceof EquipmentGraphicGanttPane.EquipmentGanttViewHandles h) {
            return h.paintFullContentForPrint();
        }
        return null;
    }

    private static Runnable extractViewportRepaint(BorderPane gantt) {
        Object ud = gantt.getUserData();
        if (ud instanceof EquipmentGraphicGanttPane.EquipmentGanttViewHandles h) {
            return h.scheduleViewportRepaint();
        }
        return null;
    }

    private static void prepareGanttForPrinting(BorderPane gantt) {
        Node bottom = gantt.getBottom();
        if (bottom != null) {
            bottom.setVisible(false);
            bottom.setManaged(false);
        }
        Object ud = gantt.getUserData();
        if (!(ud instanceof EquipmentGraphicGanttPane.EquipmentGanttViewHandles h)) {
            return;
        }
        resetScrollPaneOrigin(h.leftBodyScroll());
        resetScrollPaneOrigin(h.timelineScroll());
        resetScrollPaneOrigin(h.headerScroll());
        expandScrollPaneToFullContent(h.leftBodyScroll());
        expandScrollPaneToFullContent(h.timelineScroll());
        expandScrollPaneToFullContent(h.headerScroll());
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
        sp.setHbarPolicy(ScrollPane.ScrollBarPolicy.NEVER);
        sp.setVbarPolicy(ScrollPane.ScrollBarPolicy.NEVER);
    }

    /** snapshot 失敗時のみ。 */
    private static Parent vectorFallbackFit(
            BorderPane gantt, double paperW, double paperH, double bw, double bh) {
        double scale = Math.min(paperW / bw, paperH / bh);
        if (scale > 1.0) {
            scale = 1.0;
        }
        Group holder = new Group(gantt);
        gantt.setLayoutX(0);
        gantt.setLayoutY(0);
        if (scale < 1.0 - 1e-9) {
            holder.getTransforms().add(new Scale(scale, scale, 0, 0));
        }
        StackPane paper = new StackPane(holder);
        paper.setAlignment(Pos.TOP_LEFT);
        paper.setPrefSize(paperW, paperH);
        paper.setMinSize(paperW, paperH);
        paper.setMaxSize(paperW, paperH);
        return paper;
    }
}
