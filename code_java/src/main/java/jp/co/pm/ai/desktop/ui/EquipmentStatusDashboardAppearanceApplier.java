package jp.co.pm.ai.desktop.ui;

import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Label;
import javafx.scene.effect.DropShadow;
import javafx.scene.layout.FlowPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;

import jp.co.pm.ai.desktop.config.EquipmentStatusDashboardAppearancePrefs;

/** {@link EquipmentStatusDashboardAppearancePrefs} を FlowPane / カード Node へ反映する。 */
public final class EquipmentStatusDashboardAppearanceApplier {

    /** カード1枚ごとに生成しないよう、影の強さごとに使い回す（FX スレッドからのみ触る）。 */
    private static final Map<String, DropShadow> CARD_SHADOWS = new HashMap<>(4);

    private EquipmentStatusDashboardAppearanceApplier() {}

    /**
     * FlowPane の折返し・幅制約の算出結果。
     *
     * @param fillViewport {@code true} ならビューポート幅いっぱいに伸ばす（自動列）
     * @param wrapLength FlowPane の {@code prefWrapLength}（padding を含まない内側幅）
     * @param totalWidth 固定列時に FlowPane 自身へ与える幅。自動列では {@code -1}
     */
    public record FlowLayoutSpec(boolean fillViewport, double wrapLength, double totalWidth) {}

    /**
     * FlowPane の折返し幅を求める。JavaFX Node に触らないためテストできる。
     *
     * @param viewportWidth ScrollPane のビューポート幅（スクロールバー分は既に除かれている）。
     *     未レイアウト時は 0 以下を渡す
     */
    public static FlowLayoutSpec computeFlowSpec(
            EquipmentStatusDashboardAppearancePrefs prefs,
            boolean fullscreen,
            double viewportWidth,
            double padLeft,
            double padRight) {
        double cardW = snappedCardWidth(prefs, fullscreen);
        if (usesAutoColumnLayout(prefs)) {
            double inner =
                    viewportWidth > 1
                            ? Math.max(cardW, viewportWidth - padLeft - padRight)
                            : cardW;
            return new FlowLayoutSpec(true, inner, -1);
        }
        double wrap = fixedColumnWrapInnerWidth(prefs, fullscreen);
        return new FlowLayoutSpec(false, wrap, wrap + padLeft + padRight);
    }

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
        Insets pad = pane.getPadding();
        FlowLayoutSpec spec =
                computeFlowSpec(prefs, fullscreen, viewportWidth, pad.getLeft(), pad.getRight());
        pane.setPrefWrapLength(spec.wrapLength());
        if (spec.fillViewport()) {
            pane.setMinWidth(Region.USE_COMPUTED_SIZE);
            pane.setPrefWidth(Region.USE_COMPUTED_SIZE);
            pane.setMaxWidth(Double.MAX_VALUE);
        } else {
            pane.setPrefWidth(spec.totalWidth());
            pane.setMinWidth(spec.totalWidth());
            pane.setMaxWidth(spec.totalWidth());
        }
        return spec.fillViewport();
    }

    /** 固定列数時は ScrollPane の fitToWidth を無効にし、列幅がビューポートに潰されないようにする。 */
    public static boolean scrollShouldFitToWidth(EquipmentStatusDashboardAppearancePrefs prefs) {
        return usesAutoColumnLayout(prefs);
    }

    /** レイアウト用に切り上げたカード幅（FlowPane 折返しとカード shell の幅を揃える）。 */
    public static double snappedCardWidth(
            EquipmentStatusDashboardAppearancePrefs prefs, boolean fullscreen) {
        if (prefs == null) {
            return EquipmentStatusDashboardAppearancePrefs.defaults().effectiveCardWidth(fullscreen);
        }
        return Math.ceil(prefs.effectiveCardWidth(fullscreen));
    }

    /** 固定列数時の FlowPane 内側（padding 除く）折返し幅。自動列のときは {@code -1}。 */
    public static double fixedColumnWrapInnerWidth(
            EquipmentStatusDashboardAppearancePrefs prefs, boolean fullscreen) {
        if (prefs == null || usesAutoColumnLayout(prefs)) {
            return -1;
        }
        double cardW = snappedCardWidth(prefs, fullscreen);
        int cols = prefs.columnCount();
        double gap = Math.ceil(prefs.cardGapH());
        return cols * cardW + Math.max(0, cols - 1) * gap;
    }

    public static boolean usesAutoColumnLayout(EquipmentStatusDashboardAppearancePrefs prefs) {
        return prefs == null || prefs.columnCount() <= 0;
    }

    public static void applyFlowHostLayout(HBox host, FlowPane pane, boolean fillViewportWidth) {
        if (host == null || pane == null) {
            return;
        }
        if (fillViewportWidth) {
            host.setAlignment(Pos.TOP_LEFT);
            HBox.setHgrow(pane, Priority.ALWAYS);
        } else {
            host.setAlignment(Pos.TOP_CENTER);
            HBox.setHgrow(pane, Priority.NEVER);
        }
    }

    public static void applyCardShell(
            VBox card,
            EquipmentStatusDashboardAppearancePrefs prefs,
            boolean fullscreen) {
        if (card == null || prefs == null) {
            return;
        }
        double width = snappedCardWidth(prefs, fullscreen);
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

    public static void applyLabelFont(
            Label label, EquipmentStatusDashboardAppearancePrefs prefs, double sizePx) {
        applyLabelFont(label, prefs, sizePx, false);
    }

    /**
     * ラベルのフォントを設定する。フォント種別の指定有無で経路を分けると CSS の {@code -fx-font-size} に
     * 負けてサイズ指定が無効化されるため、常にインラインスタイルへまとめる。
     */
    public static void applyLabelFont(
            Label label,
            EquipmentStatusDashboardAppearancePrefs prefs,
            double sizePx,
            boolean bold) {
        if (label == null || prefs == null) {
            return;
        }
        StringBuilder sb = new StringBuilder(72);
        sb.append(String.format(Locale.ROOT, "-fx-font-size: %spx;", sizePx));
        if (bold) {
            sb.append("-fx-font-weight: bold;");
        }
        if (!prefs.fontFamily().isBlank()) {
            sb.append("-fx-font-family: \"")
                    .append(prefs.fontFamily().replace("\"", ""))
                    .append("\";");
        }
        label.setStyle(sb.toString());
    }

    public static void applyMachineLabel(Label label, EquipmentStatusDashboardAppearancePrefs prefs) {
        if (prefs == null) {
            return;
        }
        applyLabelFont(label, prefs, prefs.machineFontPx(), true);
    }

    private static DropShadow cardShadow(String style) {
        if (style == null || EquipmentStatusDashboardAppearancePrefs.SHADOW_NONE.equals(style)) {
            return null;
        }
        return CARD_SHADOWS.computeIfAbsent(
                style,
                key ->
                        switch (key) {
                            case EquipmentStatusDashboardAppearancePrefs.SHADOW_MEDIUM ->
                                    new DropShadow(10, 0, 3, Color.color(0, 0, 0, 0.22));
                            case EquipmentStatusDashboardAppearancePrefs.SHADOW_STRONG ->
                                    new DropShadow(16, 0, 5, Color.color(0, 0, 0, 0.32));
                            default -> new DropShadow(6, 0, 2, Color.color(0, 0, 0, 0.12));
                        });
    }
}
