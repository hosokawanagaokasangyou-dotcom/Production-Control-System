package jp.co.pm.ai.desktop.ui;

import java.util.Locale;

import javafx.scene.shape.Line;
import javafx.scene.shape.StrokeLineCap;

/** 設備ガント・担当バッジワイヤーの線種（ストロークダッシュ）。 */
public enum EquipmentGanttPersonBadgeWireDashStyle {
    SOLID("実線"),
    DASHED("破線"),
    DOTTED("点線"),
    DASH_DOT("一点鎖線");

    private final String labelJa;

    EquipmentGanttPersonBadgeWireDashStyle(String labelJa) {
        this.labelJa = labelJa;
    }

    public String labelJa() {
        return labelJa;
    }

    /** セッション JSON 等に保存するキー（enum 名と同一）。 */
    public String storedKey() {
        return name();
    }

    public static EquipmentGanttPersonBadgeWireDashStyle fromStored(String s) {
        if (s == null || s.isBlank()) {
            return SOLID;
        }
        try {
            return valueOf(s.strip().toUpperCase(Locale.ROOT));
        } catch (IllegalArgumentException e) {
            return SOLID;
        }
    }

    /**
     * @param zoom 表示倍率（ダッシュ長を帯に合わせてスケール）
     */
    public void applyTo(Line line, double zoom) {
        line.getStrokeDashArray().clear();
        double u = Math.max(1.0, zoom);
        switch (this) {
            case SOLID -> line.setStrokeLineCap(StrokeLineCap.SQUARE);
            case DASHED -> {
                line.setStrokeLineCap(StrokeLineCap.BUTT);
                line.getStrokeDashArray().addAll(8 * u, 5 * u);
            }
            case DOTTED -> {
                line.setStrokeLineCap(StrokeLineCap.ROUND);
                line.getStrokeDashArray().addAll(1.5 * u, 4 * u);
            }
            case DASH_DOT -> {
                line.setStrokeLineCap(StrokeLineCap.BUTT);
                line.getStrokeDashArray().addAll(10 * u, 4 * u, 2 * u, 4 * u);
            }
        }
    }
}
