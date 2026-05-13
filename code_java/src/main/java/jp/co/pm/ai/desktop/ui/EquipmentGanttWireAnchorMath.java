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

    /** 円環配置で隣接バッジの角度差（ラジアン）。画面で<strong>時計回り</strong>に {@code 45°}。 */
    private static final double PERSON_BADGE_RING_STEP_CLOCKWISE_RAD = Math.PI / 4.0;

    /**
     * 担当バッジをワイヤーアンカー周りの円環上に配置するための角度列（ラジアン）。
     *
     * <p>JavaFX 座標系（+X 右、+Y 下）。①最初のバッジは<strong>右上がり 45°</strong>（{@code -π/4}）。②そこから<strong>時計回りに
     * 45°</strong>（{@code π/4}）ずつ進める。バッジ数が 9 以上のとき角が一周して同じ方向に重なる場合がある。
     *
     * @param badgeCount バッジ数（正の整数想定）
     */
    public static double[] personBadgeRadialAnglesRad(int badgeCount) {
        if (badgeCount <= 0) {
            return new double[0];
        }
        double[] angles = new double[badgeCount];
        for (int i = 0; i < badgeCount; i++) {
            angles[i] =
                    PERSON_BADGE_RING_START_ANGLE_RAD + PERSON_BADGE_RING_STEP_CLOCKWISE_RAD * i;
        }
        return angles;
    }
}
