package jp.co.pm.ai.desktop.print;

import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.layout.StackPane;
import javafx.scene.paint.Color;
import javafx.print.PageLayout;

import jp.co.pm.ai.desktop.ui.EquipmentGraphicGanttPane;

/**
 * 設備ガント PDF／印刷の組み立て専用エントリ。
 *
 * <p>画面用ガントチャートは使わず、{@link EquipmentGraphicGanttPane#buildDedicatedPrintSheet} が
 * 用紙サイズに合わせた印刷専用レイアウト（横書き時刻見出し・幅フィット・ベクター）を生成する。
 */
public final class EquipmentGanttPrintCompositor {

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

        Parent sheet = EquipmentGraphicGanttPane.buildDedicatedPrintSheet(spec, paperW, paperH);
        if (sheet.getScene() == null) {
            new Scene(sheet, paperW, paperH, Color.WHITE);
        }
        sheet.applyCss();
        sheet.layout();
        return sheet;
    }
}
