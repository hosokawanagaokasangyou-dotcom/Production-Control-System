package jp.co.pm.ai.desktop.print;

import javafx.geometry.Bounds;
import javafx.geometry.Pos;
import javafx.scene.Group;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.StackPane;
import javafx.scene.transform.Scale;
import javafx.print.PageLayout;

/**
 * 設備ガントの {@link BorderPane} を、印刷 1 ジョブ内で「1 ページ」に収めるためのラッパー。
 *
 * <p>画面上のガントは {@link javafx.scene.control.ScrollPane} 内に全行分の高さを持つため、そのまま
 * {@link javafx.print.PrinterJob#printPage} すると複数物理ページに分割される。可印刷幅・高さに
 * 収まるよう等比スケールし、{@link StackPane} で用紙相当サイズに固定する。
 */
public final class EquipmentGanttPrintPageWrapper {

    private EquipmentGanttPrintPageWrapper() {}

    /**
     * レイアウト後のガント全体を測り、{@code layout} の可印刷領域に収まる最大スケール（上限 1）で縮小して中央配置する。
     *
     * @param gantt {@link jp.co.pm.ai.desktop.ui.EquipmentGraphicGanttPane#build} の戻り値
     * @param layout 用紙・余白が決まった {@link PageLayout}
     * @return 印刷に渡すルート（常に非 null）
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
        Scene measureScene = new Scene(gantt);
        java.util.Objects.requireNonNull(measureScene, "scene");
        gantt.applyCss();
        gantt.layout();
        Bounds b = gantt.getLayoutBounds();
        double w = Math.max(1.0, b.getWidth());
        double h = Math.max(1.0, b.getHeight());
        double scale = Math.min(pw / w, ph / h);
        if (scale > 1.0) {
            scale = 1.0;
        }
        Group holder = new Group(gantt);
        gantt.setLayoutX(-b.getMinX());
        gantt.setLayoutY(-b.getMinY());
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
