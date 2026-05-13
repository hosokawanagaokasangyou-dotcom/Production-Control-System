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

    /** 円環配置で最初のバッジを置く方向（ラジアン）。画面で右上がり 45°（+X と上方向の中間、{@code atan2} 基準）。 */
    private static final double PERSON_BADGE_RING_START_ANGLE_RAD = -Math.PI / 4.0;

    /**
     * 担当バッジをワイヤーアンカー周りの円環上に均等配置するための角度列（ラジアン）。
     *
     * <p>JavaFX 座標系（+X 右、+Y 下）。最初のバッジは<strong>右上がり 45°</strong>から始め、{@code 2π/n}
     * 刻みで円周へ並べる（全周等間隔）。
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
            angles[i] = PERSON_BADGE_RING_START_ANGLE_RAD + step * i;
        }
        return angles;
    }
}
