package jp.co.pm.ai.desktop.print;

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
 * 設備ガントを印刷用に「1 物理ページ」へ収める（完全ベクター）。
 *
 * <p>{@link EquipmentGraphicGanttPane#build} の {@code highQualityPrint=true} でタイムラインは
 * {@link javafx.scene.shape.Shape}／{@link javafx.scene.text.Text} になり、{@link javafx.print.PrinterJob#printPage}
 * で PDF にベクターとして出力される。ラスタ {@code snapshot} は使わない。
 */
public final class EquipmentGanttPrintPageWrapper {

    private static final double PAGE_FILL_MARGIN = 0.992;

    private EquipmentGanttPrintPageWrapper() {}

    /**
     * ガント 1 日分を用紙の可印刷領域に収めた {@link Parent} を返す。
     *
     * @param gantt {@link EquipmentGraphicGanttPane#build}（{@code highQualityPrint=true}）
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

        double cw = contentWidth(handles, gantt);
        double ch = contentHeight(handles, gantt);
        pulseLayoutForPrint(gantt, cw, ch);

        return vectorFitOnPaper(gantt, paperW, paperH, cw, ch);
    }

    private static Parent vectorFitOnPaper(
            BorderPane gantt, double paperW, double paperH, double contentW, double contentH) {
        double targetW = paperW * PAGE_FILL_MARGIN;
        double targetH = paperH * PAGE_FILL_MARGIN;
        double scale = targetW / contentW;
        double scaledH = contentH * scale;
        if (scaledH > targetH) {
            scale = targetH / contentH;
        }
        if (!Double.isFinite(scale) || scale <= 0) {
            scale = 1.0;
        }

        gantt.setLayoutX(0);
        gantt.setLayoutY(0);
        Group holder = new Group(gantt);
        holder.getTransforms().add(new Scale(scale, scale, 0, 0));

        double usedW = contentW * scale;
        double usedH = contentH * scale;
        double tx = (paperW - usedW) * 0.5;
        double ty = (paperH - usedH) * 0.5;
        holder.setTranslateX(Math.max(0, tx));
        holder.setTranslateY(Math.max(0, ty));

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

        new Scene(paper, paperW, paperH, Color.WHITE);
        paper.applyCss();
        paper.layout();
        return paper;
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
        return Math.max(1.0, gantt.getLayoutBounds().getWidth());
    }

    private static double contentHeight(
            EquipmentGraphicGanttPane.EquipmentGanttViewHandles handles, BorderPane gantt) {
        if (handles != null && handles.printContentHeight() > 1.0) {
            return handles.printContentHeight();
        }
        return Math.max(1.0, gantt.getLayoutBounds().getHeight());
    }

    private static void pulseLayoutForPrint(BorderPane gantt, double cw, double ch) {
        if (gantt.getScene() == null) {
            new Scene(gantt, cw, ch, Color.WHITE);
        }
        gantt.applyCss();
        gantt.layout();
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
            flattenHeaderRowForPrint(handles, head, headerH);
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
        expandScrollPaneToFullContent(handles.leftBodyScroll());
        expandScrollPaneToFullContent(handles.timelineScroll());

        gantt.applyCss();
        gantt.layout();
    }

    /**
     * 印刷時は {@link ScrollPane} 内の時刻見出しが PDF に出ないことがあるため、見出し行から ScrollPane を外し
     * コンテンツを直接載せる。左見出しの幅バインドも解除する。
     */
    private static void flattenHeaderRowForPrint(
            EquipmentGraphicGanttPane.EquipmentGanttViewHandles handles, HBox head, double headerH) {
        if (!head.getChildren().isEmpty()) {
            Node left = head.getChildren().get(0);
            if (left instanceof HBox leftHead) {
                try {
                    leftHead.minWidthProperty().unbind();
                } catch (RuntimeException ignored) {
                    // 未バインド
                }
                double lw = handles.printLeftWidth();
                if (lw > 0.5) {
                    leftHead.setMinWidth(lw);
                    leftHead.setPrefWidth(lw);
                    leftHead.setMaxWidth(lw);
                }
            }
        }

        HBox headerContent = handles.headerRightContent();
        double headerRightW =
                Math.max(
                        1.0,
                        handles.printContentWidth() > 0.5 && handles.printLeftWidth() > 0.5
                                ? handles.printContentWidth() - handles.printLeftWidth()
                                : handles.printTimelineWidth());
        if (headerContent != null) {
            headerContent.setMinHeight(headerH);
            headerContent.setPrefHeight(headerH);
            headerContent.setMaxHeight(headerH);
            headerContent.setMinWidth(headerRightW);
            headerContent.setPrefWidth(headerRightW);
            headerContent.setMaxWidth(headerRightW);
            headerContent.setVisible(true);
            headerContent.setManaged(true);
        }

        if (head.getChildren().size() >= 2 && head.getChildren().get(1) instanceof ScrollPane headerScroll) {
            Node content = headerScroll.getContent();
            if (content != null) {
                head.getChildren().set(1, content);
            } else if (headerContent != null) {
                head.getChildren().set(1, headerContent);
            }
        }

        if (!head.getChildren().isEmpty()) {
            Node right = head.getChildren().get(Math.min(1, head.getChildren().size() - 1));
            HBox.setHgrow(right, Priority.ALWAYS);
            if (right instanceof Region r) {
                r.setMinHeight(headerH);
                r.setPrefHeight(headerH);
                r.setMaxHeight(headerH);
                if (headerRightW > 0.5) {
                    r.setMinWidth(headerRightW);
                    r.setPrefWidth(headerRightW);
                    r.setMaxWidth(headerRightW);
                }
            }
        }
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
