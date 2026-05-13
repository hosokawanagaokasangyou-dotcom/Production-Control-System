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

    /**
     * 担当バッジをワイヤーアンカー周りの円環上に均等配置するための角度列（ラジアン）。
     *
     * <p>JavaFX 座標系（+X 右、+Y 下）。{@code -π/2} は画面上方向。最初のバッジを上に置き、時計回りに {@code 2π/n}
     * 刻みで円周へ並べる（扇形ではなく全周の等間隔）。
     *
     * @param badgeCount バッジ数（正の整数想定）
     */
    public static double[] personBadgeRadialAnglesRad(int badgeCount) {
        if (badgeCount <= 0) {
            return new double[0];
        }
        double[] angles = new double[badgeCount];
        double step = (2.0 * Math.PI) / badgeCount;
        for (int i = 0; i < badgeCount; i++) {
            angles[i] = -Math.PI / 2.0 + step * i;
        }
        return angles;
    }
}
