package jp.co.pm.ai.desktop.ui;

import javafx.geometry.Insets;
import javafx.scene.control.Label;
import javafx.scene.effect.DropShadow;
import javafx.scene.layout.StackPane;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;

import jp.co.pm.ai.desktop.config.PersonBadgeStyle;

/** 担当バッジ用の小型 {@link StackPane} を組み立てる。 */
public final class PersonBadgeNodeFactory {

    private PersonBadgeNodeFactory() {}

    public static StackPane createBadge(String text, PersonBadgeStyle st, double zoom, double rowLabelFontPx) {
        String t = text != null ? text : "";
        double pct = st.fontPercent() > 0 ? st.fontPercent() / 100.0 : 0.85;
        double fs = Math.max(6, rowLabelFontPx * pct * 0.9);
        Label lb = new Label(t);
        String fam = st.fontFamily();
        lb.setFont(
                fam != null && !fam.isBlank()
                        ? Font.font(fam, fs)
                        : Font.font(fs));
        lb.setTextFill(safeColor(st.textHex(), Color.WHITE));

        double padX = Math.max(3, 5 * zoom);
        double padY = Math.max(1, 2 * zoom);
        StackPane sp = new StackPane(lb);
        sp.setPadding(new Insets(padY, padX, padY, padX));

        String rad = st.pill() ? "999" : String.valueOf(Math.max(0, st.cornerRadius()));
        double sw = Math.max(0, st.strokeWidth());
        String borderPart =
                sw > 0.01
                        ? String.format(
                                "-fx-border-color: %s; -fx-border-radius: %s; -fx-border-width: %.2fpx;",
                                esc(st.strokeHex(), "#1e40af"), rad, sw)
                        : "";
        sp.setStyle(
                String.format(
                        "-fx-background-color: %s; -fx-background-radius: %s; %s",
                        esc(st.fillHex(), "#2563eb"), rad, borderPart));

        DropShadow glow = new DropShadow();
        glow.setColor(safeColor(st.glowColorHex(), Color.web("#38bdf8")));
        glow.setRadius(Math.max(0, st.glowRadius()));
        double spRead = st.glowSpread();
        if (spRead >= 0 && spRead <= 1) {
            glow.setSpread(spRead);
        }
        sp.setEffect(glow);

        double op = st.opacity();
        sp.setOpacity(
                Math.max(0.0, Math.min(1.0, Double.isFinite(op) ? op : 1.0)));

        return sp;
    }

    private static String esc(String hex, String fallback) {
        String h = hex != null ? hex.strip() : "";
        if (h.isEmpty()) {
            return fallback;
        }
        try {
            Color.web(h);
            return h;
        } catch (IllegalArgumentException e) {
            return fallback;
        }
    }

    private static Color safeColor(String hex, Color fallback) {
        if (hex == null || hex.isBlank()) {
            return fallback;
        }
        try {
            return Color.web(hex.trim());
        } catch (IllegalArgumentException e) {
            return fallback;
        }
    }
}
