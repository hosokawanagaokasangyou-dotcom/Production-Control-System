package jp.co.pm.ai.desktop.print;

import javafx.geometry.Bounds;
import javafx.geometry.Pos;
import javafx.scene.Group;
import javafx.scene.Node;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
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
 * <p>主経路は {@link javafx.print.PrinterJob#printPage} 向けのベクター配置（{@link Group}＋{@link Scale}）。
 * 左列・ラベル・担当バッジはプリンタ解像度で描画される。タイムライン帯は {@link javafx.scene.canvas.Canvas}
 * のため高解像度 build（{@link EquipmentGraphicGanttPane#PRINT_LAYOUT_SCALE}）と組み合わせる。
 */
public final class EquipmentGanttPrintPageWrapper {

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
        pulseLayoutForPrint(gantt, handles, cw, ch);

        return vectorFitCenteredOnPaper(gantt, paperW, paperH, cw, ch);
    }

    private static EquipmentGraphicGanttPane.EquipmentGanttViewHandles viewHandles(BorderPane gantt) {
        Object ud = gantt.getUserData();
        if (ud instanceof EquipmentGraphicGanttPane.EquipmentGanttViewHandles h) {
            return h;
        }
        return null;
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
            double cw,
            double ch) {
        Scene scene = new Scene(gantt, cw, ch, Color.WHITE);
        java.util.Objects.requireNonNull(scene, "scene");
        gantt.applyCss();
        gantt.layout();
        runFullTimelinePaint(handles);
        gantt.applyCss();
        gantt.layout();
    }

    /**
     * シーン上でレイアウト済みのガントを、用紙座標系へ等比スケールして中央配置する（ラスタ中間なし）。
     */
    private static Parent vectorFitCenteredOnPaper(
            BorderPane gantt, double paperW, double paperH, double contentW, double contentH) {
        double scale = Math.min(paperW / contentW, paperH / contentH);
        if (!Double.isFinite(scale) || scale <= 0) {
            scale = 1.0;
        }

        gantt.setLayoutX(0);
        gantt.setLayoutY(0);
        Group holder = new Group(gantt);
        holder.getTransforms().add(new Scale(scale, scale, 0, 0));

        double dw = contentW * scale;
        double dh = contentH * scale;
        holder.setTranslateX((paperW - dw) * 0.5);
        holder.setTranslateY((paperH - dh) * 0.5);

        Region bg = new Region();
        bg.setMinSize(paperW, paperH);
        bg.setPrefSize(paperW, paperH);
        bg.setMaxSize(paperW, paperH);
        bg.setStyle("-fx-background-color: white;");

        StackPane paper = new StackPane(bg, holder);
        paper.setAlignment(Pos.CENTER);
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
}
