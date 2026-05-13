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
     * 担当バッジをワイヤーアンカー周りに均等配置するための角度列（ラジアン）。
     *
     * <p>JavaFX 座標系（+X 右、+Y 下）。{@code -π/2} は画面上方向。バーを跨がないよう主に上側の扇形に収める。
     *
     * @param badgeCount バッジ数（正の整数想定）
     */
    public static double[] personBadgeRadialAnglesRad(int badgeCount) {
        if (badgeCount <= 0) {
            return new double[0];
        }
        if (badgeCount == 1) {
            return new double[] {-Math.PI / 2.0};
        }
        double minTotal = Math.toRadians(28.0) * (badgeCount - 1);
        double desired = Math.toRadians(40.0 + 22.0 * (badgeCount - 1));
        double maxTotal = Math.toRadians(210.0);
        double totalArc = Math.min(maxTotal, Math.max(minTotal, desired));
        double center = -Math.PI / 2.0;
        double half = totalArc / 2.0;
        double start = center - half;
        double[] angles = new double[badgeCount];
        for (int i = 0; i < badgeCount; i++) {
            angles[i] = start + totalArc * i / (badgeCount - 1);
        }
        return angles;
    }
}
