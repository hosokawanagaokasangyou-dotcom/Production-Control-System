package jp.co.pm.ai.desktop.ui;

import java.util.Locale;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.chart.PieChart;
import javafx.scene.control.Label;
import javafx.scene.effect.DropShadow;
import javafx.scene.layout.FlowPane;
import javafx.scene.layout.Region;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;
import javafx.scene.text.FontPosture;
import javafx.scene.text.FontWeight;

import jp.co.pm.ai.desktop.config.EquipmentStatusDashboardAppearancePrefs;

/** {@link EquipmentStatusDashboardAppearancePrefs} を FlowPane / カード Node へ反映する。 */
public final class EquipmentStatusDashboardAppearanceApplier {

    private EquipmentStatusDashboardAppearanceApplier() {}

    /**
     * FlowPane の折返し・幅制約を更新する。
     *
     * @return {@code true} ならビューポート幅いっぱいに FlowPane を伸ばす（自動列）。{@code false} なら固定列幅で中央寄せ。
     */
    public static boolean configureFlowPane(
            FlowPane pane,
            EquipmentStatusDashboardAppearancePrefs prefs,
            boolean fullscreen,
            double viewportWidth) {
        if (pane == null || prefs == null) {
            return true;
        }
        pane.setHgap(prefs.cardGapH());
        pane.setVgap(prefs.cardGapV());
        if (usesAutoColumnLayout(prefs)) {
            clearFixedColumnWidthConstraints(pane);
            double cardW = prefs.effectiveCardWidth(fullscreen);
            if (viewportWidth > 1) {
                Insets pad = pane.getPadding();
                double inner =
                        Math.max(
                                cardW,
                                viewportWidth - pad.getLeft() - pad.getRight() - 24);
                pane.setPrefWrapLength(inner);
            } else {
                pane.setPrefWrapLength(RegionFallback.COMPUTED);
            }
            return true;
        }
        double wrap = fixedColumnWrapInnerWidth(prefs, fullscreen);
        Insets pad = pane.getPadding();
        double totalW = wrap + pad.getLeft() + pad.getRight();
        pane.setPrefWrapLength(wrap);
        pane.setPrefWidth(totalW);
        pane.setMinWidth(totalW);
        pane.setMaxWidth(totalW);
        return false;
    }

    /** 固定列数時の FlowPane 内側（padding 除く）折返し幅。自動列のときは {@code -1}。 */
    public static double fixedColumnWrapInnerWidth(
            EquipmentStatusDashboardAppearancePrefs prefs, boolean fullscreen) {
        if (prefs == null || usesAutoColumnLayout(prefs)) {
            return -1;
        }
        double cardW = prefs.effectiveCardWidth(fullscreen);
        int cols = prefs.columnCount();
        return cols * cardW + Math.max(0, cols - 1) * prefs.cardGapH();
    }

    public static boolean usesAutoColumnLayout(EquipmentStatusDashboardAppearancePrefs prefs) {
        return prefs == null || prefs.columnCount() <= 0;
    }

    public static void applyFlowHostLayout(
            javafx.scene.layout.HBox host, FlowPane pane, boolean fillViewportWidth) {
        if (host == null || pane == null) {
            return;
        }
        if (fillViewportWidth) {
            host.setAlignment(Pos.TOP_LEFT);
            javafx.scene.layout.HBox.setHgrow(pane, javafx.scene.layout.Priority.ALWAYS);
        } else {
            host.setAlignment(Pos.TOP_CENTER);
            javafx.scene.layout.HBox.setHgrow(pane, javafx.scene.layout.Priority.NEVER);
        }
    }

    private static void clearFixedColumnWidthConstraints(FlowPane pane) {
        pane.setMinWidth(Region.USE_COMPUTED_SIZE);
        pane.setPrefWidth(Region.USE_COMPUTED_SIZE);
        pane.setMaxWidth(Double.MAX_VALUE);
    }

    public static void applyCardShell(
            VBox card,
            EquipmentStatusDashboardAppearancePrefs prefs,
            boolean fullscreen) {
        if (card == null || prefs == null) {
            return;
        }
        double width = prefs.effectiveCardWidth(fullscreen);
        card.setPadding(new Insets(prefs.cardPadding()));
        card.setPrefWidth(width);
        card.setMinWidth(width);
        card.setMaxWidth(width);
        card.setStyle(
                String.format(
                        Locale.ROOT,
                        "-fx-background-radius: %spx; -fx-border-radius: %spx;",
                        prefs.cardBorderRadius(),
                        prefs.cardBorderRadius()));
        card.setEffect(cardShadow(prefs.cardShadowStyle()));
    }

    public static void applyLabelFont(Label label, EquipmentStatusDashboardAppearancePrefs prefs, double sizePx) {
        if (label == null || prefs == null) {
            return;
        }
        if (prefs.fontFamily().isBlank()) {
            label.setStyle(String.format(Locale.ROOT, "-fx-font-size: %spx;", sizePx));
        } else {
            label.setFont(Font.font(prefs.fontFamily(), FontWeight.NORMAL, sizePx));
        }
    }

    public static void applyMachineLabel(Label label, EquipmentStatusDashboardAppearancePrefs prefs) {
        if (label == null || prefs == null) {
            return;
        }
        if (prefs.fontFamily().isBlank()) {
            label.setStyle(
                    String.format(
                            Locale.ROOT,
                            "-fx-font-size: %spx; -fx-font-weight: bold;",
                            prefs.machineFontPx()));
        } else {
            label.setFont(
                    Font.font(
                            prefs.fontFamily(),
                            FontWeight.BOLD,
                            FontPosture.REGULAR,
                            prefs.machineFontPx()));
        }
    }

    public static StackPane buildPieChart(
            double completionPct, EquipmentStatusDashboardAppearancePrefs prefs) {
        double done = Math.max(0.0, Math.min(100.0, completionPct));
        double remain = 100.0 - done;
        PieChart chart =
                new PieChart(
                        javafx.collections.FXCollections.observableArrayList(
                                new PieChart.Data("完了", done),
                                new PieChart.Data("残り", remain > 0 ? remain : 0.01)));
        chart.setAnimated(false);
        chart.setLegendVisible(false);
        chart.setLabelsVisible(false);
        double sz = prefs != null ? prefs.chartSizePx() : 96;
        chart.setPrefSize(sz, sz);
        chart.setMinSize(sz, sz);
        chart.setMaxSize(sz, sz);
        chart.getStyleClass().add("pm-equipment-status-pie");

        Label pct = new Label(String.format(Locale.ROOT, "%.0f%%", done));
        pct.getStyleClass().add("pm-equipment-status-pct-label");
        if (prefs != null) {
            applyLabelFont(pct, prefs, prefs.pctFontPx());
        }

        StackPane pane = new StackPane(chart, pct);
        pane.setAlignment(javafx.geometry.Pos.CENTER);
        if (prefs != null) {
            stylePieChart(chart, prefs);
            if (prefs.chartShadowEnabled()) {
                DropShadow ds = new DropShadow(8, 0, 2, Color.color(0, 0, 0, 0.25));
                chart.setEffect(ds);
            }
        }
        return pane;
    }

    private static void stylePieChart(PieChart chart, EquipmentStatusDashboardAppearancePrefs prefs) {
        Runnable apply =
                () -> {
                    for (PieChart.Data d : chart.getData()) {
                        if (d.getNode() == null) {
                            continue;
                        }
                        String color =
                                "完了".equals(d.getName())
                                        ? prefs.chartDoneColorHex()
                                        : prefs.chartRemainColorHex();
                        if (EquipmentStatusDashboardAppearancePrefs.CHART_DEPTH.equals(prefs.chartStyle())) {
                            String lighter = lightenHex(color, 0.18);
                            d.getNode()
                                    .setStyle(
                                            String.format(
                                                    Locale.ROOT,
                                                    "-fx-pie-color: linear-gradient(to bottom, %s, %s);",
                                                    lighter,
                                                    color));
                        } else {
                            d.getNode()
                                    .setStyle(
                                            String.format(
                                                    Locale.ROOT, "-fx-pie-color: %s;", color));
                        }
                    }
                };
        chart.getData().forEach(d -> d.nodeProperty().addListener((o, a, n) -> apply.run()));
        javafx.application.Platform.runLater(apply);
    }

    private static DropShadow cardShadow(String style) {
        return switch (style != null ? style : EquipmentStatusDashboardAppearancePrefs.SHADOW_SUBTLE) {
            case EquipmentStatusDashboardAppearancePrefs.SHADOW_NONE -> null;
            case EquipmentStatusDashboardAppearancePrefs.SHADOW_MEDIUM ->
                    new DropShadow(10, 0, 3, Color.color(0, 0, 0, 0.22));
            case EquipmentStatusDashboardAppearancePrefs.SHADOW_STRONG ->
                    new DropShadow(16, 0, 5, Color.color(0, 0, 0, 0.32));
            default -> new DropShadow(6, 0, 2, Color.color(0, 0, 0, 0.12));
        };
    }

    private static String lightenHex(String hex, double amount) {
        try {
            Color c = Color.web(hex);
            return String.format(
                    Locale.ROOT,
                    "#%02x%02x%02x",
                    (int) Math.min(255, c.getRed() * 255 + 255 * amount),
                    (int) Math.min(255, c.getGreen() * 255 + 255 * amount),
                    (int) Math.min(255, c.getBlue() * 255 + 255 * amount));
        } catch (Exception ex) {
            return hex;
        }
    }

    /** {@code FlowPane#setPrefWrapLength} 用（自動列）。 */
    private static final class RegionFallback {
        static final double COMPUTED = Region.USE_COMPUTED_SIZE;

        private RegionFallback() {}
    }
}
