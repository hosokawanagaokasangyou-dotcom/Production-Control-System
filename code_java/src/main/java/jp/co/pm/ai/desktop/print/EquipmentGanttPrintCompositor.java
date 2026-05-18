package jp.co.pm.ai.desktop.print;

import javafx.geometry.Pos;
import javafx.scene.Group;
import javafx.scene.Node;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.GridPane;
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
 * 設備ガント PDF／印刷の組み立て専用エントリ。
 *
 * <p>手順は次の 3 段のみ。
 *
 * <ol>
 *   <li>{@linkplain #resolveSlotWidthPercent 用紙幅から時刻スロット幅％を決定}（左 3 列・進捗列は維持、時刻軸だけ伸縮）
 *   <li>{@link EquipmentGraphicGanttPane#composePrintPage} で ScrollPane なしのベクターチャートを生成
 *   <li>{@linkplain #mountOnPaper 用紙サイズの Parent} に載せて {@code printPage} へ渡す
 * </ol>
 */
public final class EquipmentGanttPrintCompositor {

    /** 可印刷幅に対するチャート幅（左右の極小余白のみ）。 */
    static final double WIDTH_FILL = 0.995;

    /** 高さがはみ出すときだけ縮小する上限。 */
    static final double HEIGHT_FILL = 0.995;

    private EquipmentGanttPrintCompositor() {}

  /**
   * 1 暦日分の印刷ルートノードを組み立てる。
   *
   * @return 用紙可印刷領域と同じサイズの {@link Parent}（白背景・ベクター）
   */
    public static Parent composePage(EquipmentGanttPrintPageSpec spec, PageLayout layout) {
        if (spec == null || layout == null) {
            return new StackPane();
        }
        double paperW = layout.getPrintableWidth();
        double paperH = layout.getPrintableHeight();
        if (!Double.isFinite(paperW) || !Double.isFinite(paperH) || paperW < 2 || paperH < 2) {
            return new StackPane();
        }

        double slotPct = resolveSlotWidthPercent(spec, paperW);
        Parent chart = EquipmentGraphicGanttPane.composePrintPage(spec, slotPct);
        return mountOnPaper(chart, paperW, paperH);
    }

    /**
     * プローブ（画面と同じスロット％）で左列・進捗・時刻軸の幅を測り、時刻軸だけを {@link #WIDTH_FILL} まで広げる％を返す。
     */
    static double resolveSlotWidthPercent(EquipmentGanttPrintPageSpec spec, double printableWidthPx) {
        EquipmentGraphicGanttPane.EquipmentGanttPrintMetrics m =
                EquipmentGraphicGanttPane.measurePrintLayout(spec);
        double targetW = printableWidthPx * WIDTH_FILL;
        double nonTimeline = Math.max(0d, m.contentWidth() - m.timelineWidth());
        double targetTimeline = targetW - nonTimeline;
        if (m.timelineWidth() < 2 || targetTimeline < 2) {
            return spec.slotWidthPercent();
        }
        return Math.clamp(spec.slotWidthPercent() * (targetTimeline / m.timelineWidth()), 30d, 2000d);
    }

    private static Parent mountOnPaper(Parent chart, double paperW, double paperH) {
        double cw = Math.max(1.0, chart.prefWidth(-1));
        double ch = Math.max(1.0, chart.prefHeight(-1));
        if (cw <= 1 || ch <= 1) {
            cw = Math.max(cw, chart.getLayoutBounds().getWidth());
            ch = Math.max(ch, chart.getLayoutBounds().getHeight());
        }

        double scale = 1.0;
        double maxH = paperH * HEIGHT_FILL;
        if (ch > maxH && ch > 1) {
            scale = maxH / ch;
        }
        double scaledW = cw * scale;
        double scaledH = ch * scale;

        chart.setLayoutX(0);
        chart.setLayoutY(0);
        Group holder = new Group(chart);
        if (Math.abs(scale - 1.0) > 1e-6) {
            holder.getTransforms().add(new Scale(scale, scale, 0, 0));
        }

        Region bg = new Region();
        bg.setMinSize(paperW, paperH);
        bg.setPrefSize(paperW, paperH);
        bg.setMaxSize(paperW, paperH);
        bg.setStyle("-fx-background-color: white;");

        StackPane paper = new StackPane(bg, holder);
        paper.setAlignment(Pos.TOP_LEFT);
        paper.setPrefSize(paperW, paperH);
        paper.setMinSize(paperW, paperH);
        paper.setMaxSize(paperW, paperH);

        holder.setTranslateX(Math.max(0, (paperW - scaledW) * 0.5));
        holder.setTranslateY(Math.max(0, (paperH - scaledH) * 0.02));

        if (chart.getScene() == null) {
            new Scene(paper, paperW, paperH, Color.WHITE);
        }
        paper.applyCss();
        paper.layout();
        return paper;
    }
}
