package jp.co.pm.ai.desktop.ui;

import java.util.Locale;

import javafx.geometry.Pos;
import javafx.scene.canvas.Canvas;
import javafx.scene.canvas.GraphicsContext;
import javafx.scene.control.Label;
import javafx.scene.effect.DropShadow;
import javafx.scene.layout.StackPane;
import javafx.scene.paint.Color;
import javafx.scene.paint.CycleMethod;
import javafx.scene.paint.LinearGradient;
import javafx.scene.paint.Paint;
import javafx.scene.paint.Stop;
import javafx.scene.shape.ArcType;
import javafx.scene.shape.StrokeLineCap;

import jp.co.pm.ai.desktop.config.EquipmentStatusDashboardAppearancePrefs;

/**
 * 達成率を表すリング。{@code Canvas} 1枚 + ラベル1枚で描くため、カードごとに {@code PieChart}
 * （凡例・ラベル・スライス Node を伴う）を生成する必要がない。
 */
public final class EquipmentStatusProgressRing extends StackPane {

    private static final DropShadow RING_SHADOW = new DropShadow(8, 0, 2, Color.color(0, 0, 0, 0.25));
    private static final Color FALLBACK_DONE = Color.web("#0d9488");
    private static final Color FALLBACK_REMAIN = Color.web("#e2e8f0");

    /** リング太さの直径に対する比率。 */
    private static final double THICKNESS_RATIO = 0.18;

    private final Canvas canvas = new Canvas();
    private final Label pctLabel = new Label();

    public EquipmentStatusProgressRing() {
        getStyleClass().add("pm-equipment-status-ring");
        pctLabel.getStyleClass().add("pm-equipment-status-pct-label");
        setAlignment(Pos.CENTER);
        getChildren().addAll(canvas, pctLabel);
    }

    /** 達成率（%）と見た目設定を反映する。100% を超える値はラベルにそのまま出し、リングは満円で止める。 */
    public void update(double completionPct, EquipmentStatusDashboardAppearancePrefs prefs) {
        EquipmentStatusDashboardAppearancePrefs p =
                prefs != null ? prefs : EquipmentStatusDashboardAppearancePrefs.defaults();
        double size = Math.max(24.0, p.chartSizePx());
        double thickness = Math.max(4.0, size * THICKNESS_RATIO);
        double inset = thickness / 2.0;
        double box = size - thickness;
        double pct = Double.isFinite(completionPct) ? Math.max(0.0, completionPct) : 0.0;
        double drawnPct = Math.min(100.0, pct);

        canvas.setWidth(size);
        canvas.setHeight(size);
        setPrefSize(size, size);
        setMinSize(size, size);
        setMaxSize(size, size);

        GraphicsContext g = canvas.getGraphicsContext2D();
        g.clearRect(0, 0, size, size);
        g.setLineWidth(thickness);
        g.setLineCap(StrokeLineCap.BUTT);
        g.setStroke(webColor(p.chartRemainColorHex(), FALLBACK_REMAIN));
        g.strokeArc(inset, inset, box, box, 90, 360, ArcType.OPEN);
        if (drawnPct > 0) {
            g.setStroke(donePaint(p));
            g.strokeArc(inset, inset, box, box, 90, -360.0 * drawnPct / 100.0, ArcType.OPEN);
        }

        pctLabel.setText(String.format(Locale.ROOT, "%.0f%%", pct));
        EquipmentStatusDashboardAppearanceApplier.applyLabelFont(pctLabel, p, p.pctFontPx(), true);
        canvas.setEffect(p.chartShadowEnabled() ? RING_SHADOW : null);
        setAccessibleText("アラジン達成率 " + Math.round(pct) + "パーセント");
    }

    private static Paint donePaint(EquipmentStatusDashboardAppearancePrefs prefs) {
        Color done = webColor(prefs.chartDoneColorHex(), FALLBACK_DONE);
        if (!EquipmentStatusDashboardAppearancePrefs.CHART_DEPTH.equals(prefs.chartStyle())) {
            return done;
        }
        return new LinearGradient(
                0,
                0,
                0,
                1,
                true,
                CycleMethod.NO_CYCLE,
                new Stop(0, lighten(done, 0.18)),
                new Stop(1, done));
    }

    static Color lighten(Color base, double amount) {
        if (base == null) {
            return FALLBACK_DONE;
        }
        double a = Math.max(0.0, Math.min(1.0, amount));
        return new Color(
                base.getRed() + (1.0 - base.getRed()) * a,
                base.getGreen() + (1.0 - base.getGreen()) * a,
                base.getBlue() + (1.0 - base.getBlue()) * a,
                base.getOpacity());
    }

    static Color webColor(String hex, Color fallback) {
        if (hex == null || hex.isBlank()) {
            return fallback;
        }
        try {
            return Color.web(hex.strip());
        } catch (IllegalArgumentException ex) {
            return fallback;
        }
    }
}
