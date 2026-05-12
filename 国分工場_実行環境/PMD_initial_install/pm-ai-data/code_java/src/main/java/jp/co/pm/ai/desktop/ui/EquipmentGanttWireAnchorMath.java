package jp.co.pm.ai.desktop.ui;

/**
 * 設備ガントのチャートバー中心（バッジワイヤーのアンカー）。{@link EquipmentGraphicGanttPane} の
 * {@code fillBar} / 担当バッジオーバーレイと同一座標系。
 */
public final class EquipmentGanttWireAnchorMath {

    private EquipmentGanttWireAnchorMath() {}

    /** 塗りつぶし矩形の水平中心（オーバーレイ X）。 */
    public static double barAnchorCenterX(double slotWidth, double zoom, int fromSlot, int toSlot) {
        double inset = 0.5 * zoom;
        double spanW = (toSlot - fromSlot + 1) * slotWidth;
        return fromSlot * slotWidth + inset + (spanW - 2 * inset) / 2;
    }

    /** チャートバー帯の垂直中心（オーバーレイ Y）。 */
    public static double barAnchorCenterY(double timelineOuterPad, double rowHeight, double zoom) {
        double innerBarTop = 3 * zoom;
        double barTop = timelineOuterPad + innerBarTop;
        double barH = rowHeight - 2 * innerBarTop;
        return barTop + barH / 2;
    }
}
