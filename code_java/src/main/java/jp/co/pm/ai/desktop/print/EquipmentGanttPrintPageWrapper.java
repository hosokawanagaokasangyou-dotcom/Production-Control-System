package jp.co.pm.ai.desktop.print;

import javafx.geometry.Pos;
import javafx.scene.Group;
import javafx.scene.Node;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.SnapshotParameters;
import javafx.scene.control.ScrollPane;
import javafx.scene.image.ImageView;
import javafx.scene.image.WritableImage;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.transform.Scale;
import javafx.print.PageLayout;

import jp.co.pm.ai.desktop.ui.EquipmentGraphicGanttPane;

/**
 * 設備ガントを印刷用に「1 物理ページ」へ収める。
 *
 * <p>画面上は {@link ScrollPane} のビューポート＋部分描画のため、そのまま {@link javafx.print.PrinterJob#printPage}
 * すると座標が崩れバッジだけが浮く等になる。印刷直前にビューポートを内容全体に広げ全再描画し、
 * ラスタ（{@code snapshot}）として用紙可印刷領域に収めてから印刷する。
 */
public final class EquipmentGanttPrintPageWrapper {

    /** スナップショット 1 辺の安全上限（GPU／Prism の上限に合わせる） */
    private static final double SNAPSHOT_MAX_EDGE = 8192;

    private EquipmentGanttPrintPageWrapper() {}

    /**
     * ガント 1 日分を A3 横の可印刷領域に収めた {@link Parent} を返す。
     *
     * @param gantt {@link EquipmentGraphicGanttPane#build} の戻り値
     * @param layout 用紙・向きが確定した {@link PageLayout}
     */
    public static Parent fitGanttToSinglePrintablePage(BorderPane gantt, PageLayout layout) {
        if (gantt == null || layout == null) {
            return new StackPane();
        }
        double pw = layout.getPrintableWidth();
        double ph = layout.getPrintableHeight();
        if (!Double.isFinite(pw) || !Double.isFinite(ph) || pw < 2 || ph < 2) {
            return gantt;
        }

        Scene measureScene = new Scene(gantt, 4096, 8192, Color.WHITE);
        java.util.Objects.requireNonNull(measureScene, "scene");
        prepareGanttForPrinting(gantt);
        gantt.applyCss();
        gantt.layout();

        Runnable repaint = extractRepaint(gantt);
        if (repaint != null) {
            repaint.run();
        }
        gantt.applyCss();
        gantt.layout();

        double preScale = 1.0;
        double bw = Math.max(1.0, gantt.getLayoutBounds().getWidth());
        double bh = Math.max(1.0, gantt.getLayoutBounds().getHeight());
        if (bw > SNAPSHOT_MAX_EDGE || bh > SNAPSHOT_MAX_EDGE) {
            preScale = Math.min(SNAPSHOT_MAX_EDGE / bw, SNAPSHOT_MAX_EDGE / bh);
        }

        SnapshotParameters snapParams = new SnapshotParameters();
        snapParams.setFill(Color.WHITE);
        if (preScale < 1.0 - 1e-9) {
            snapParams.setTransform(new Scale(preScale, preScale, 0, 0));
        }

        WritableImage img;
        try {
            img = gantt.snapshot(snapParams, null);
        } catch (RuntimeException ex) {
            return vectorFallbackFit(gantt, pw, ph, bw, bh);
        }
        if (img == null || img.getWidth() < 1 || img.getHeight() < 1) {
            return vectorFallbackFit(gantt, pw, ph, bw, bh);
        }

        double iw = img.getWidth();
        double ih = img.getHeight();
        ImageView iv = new ImageView(img);
        iv.setSmooth(true);
        iv.setPreserveRatio(true);
        StackPane paper = new StackPane(iv);
        paper.setPrefSize(pw, ph);
        paper.setMinSize(pw, ph);
        paper.setMaxSize(pw, ph);
        double fit = Math.min(pw / iw, ph / ih);
        iv.setFitWidth(iw * fit);
        iv.setFitHeight(ih * fit);
        StackPane.setAlignment(iv, Pos.CENTER);
        return paper;
    }

    private static Runnable extractRepaint(BorderPane gantt) {
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
        expandScrollPaneToFullContent(h.leftBodyScroll());
        expandScrollPaneToFullContent(h.timelineScroll());
        expandScrollPaneToFullContent(findHeaderScrollFromRoot(gantt));
    }

    private static ScrollPane findHeaderScrollFromRoot(BorderPane gantt) {
        Node center = gantt.getCenter();
        if (!(center instanceof VBox vb)) {
            return null;
        }
        for (Node ch : vb.getChildren()) {
            if (ch instanceof HBox hb) {
                for (Node c2 : hb.getChildren()) {
                    if (c2 instanceof ScrollPane sp) {
                        return sp;
                    }
                }
            }
        }
        return null;
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

    /** snapshot 失敗時のみ。可能な限り {@link StackPane}＋等比スケールで収める。 */
    private static Parent vectorFallbackFit(
            BorderPane gantt, double pw, double ph, double bw, double bh) {
        double scale = Math.min(pw / bw, ph / bh);
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
        paper.setAlignment(Pos.CENTER);
        paper.setPrefSize(pw, ph);
        paper.setMinSize(pw, ph);
        paper.setMaxSize(pw, ph);
        return paper;
    }
}
