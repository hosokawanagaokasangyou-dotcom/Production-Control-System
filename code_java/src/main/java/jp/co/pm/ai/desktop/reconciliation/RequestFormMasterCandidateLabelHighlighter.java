package jp.co.pm.ai.desktop.reconciliation;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Locale;

import javafx.geometry.Insets;
import javafx.scene.Node;
import javafx.scene.Scene;
import javafx.scene.layout.Region;
import javafx.scene.paint.Color;
import javafx.scene.text.Text;
import javafx.scene.text.TextFlow;

/** マスタ候補コンボ表示文字列で、フィルタ語と一致する部分を黄色で強調する。 */
final class RequestFormMasterCandidateLabelHighlighter {

    /** midnight-blue 既定の -fx-control-inner-background */
    private static final Color DEFAULT_LIST_BACKGROUND = Color.web("#1E3A5F");

    private static final Color HIGHLIGHT_BACKGROUND = Color.web("#FFEB3B");

    private RequestFormMasterCandidateLabelHighlighter() {}

    static Node buildGraphic(String label, Collection<String> filterKeywords, Node styleAnchor) {
        if (label == null || label.isEmpty()) {
            return new Text("");
        }
        Color listBg = resolveListBackground(styleAnchor);
        String plainStyle = textFillStyle(complementaryContrastText(listBg));
        String highlightStyle =
                textFillStyle(complementaryContrastText(HIGHLIGHT_BACKGROUND))
                        + " -fx-background-color: #FFEB3B;";

        List<String> keywords = nonEmptyKeywords(filterKeywords);
        if (keywords.isEmpty()) {
            Text plain = new Text(label);
            plain.setStyle(plainStyle);
            return plain;
        }
        boolean[] highlight = highlightMask(label, keywords);
        TextFlow flow = new TextFlow();
        flow.getStyleClass().add("request-form-master-candidate-flow");
        flow.setPadding(new Insets(1, 4, 1, 2));
        int i = 0;
        while (i < label.length()) {
            boolean on = highlight[i];
            int j = i + 1;
            while (j < label.length() && highlight[j] == on) {
                j++;
            }
            Text segment = new Text(label.substring(i, j));
            segment.setStyle(on ? highlightStyle : plainStyle);
            segment.getStyleClass().add(on ? "request-form-master-candidate-highlight" : "request-form-master-candidate-plain");
            flow.getChildren().add(segment);
            i = j;
        }
        return flow;
    }

    /** 背景色の補色（色相 +180°）をベースに、明度で読みやすい文字色を返す。 */
    static Color complementaryContrastText(Color background) {
        if (background.getSaturation() < 0.12) {
            double lum =
                    0.2126 * background.getRed()
                            + 0.7152 * background.getGreen()
                            + 0.0722 * background.getBlue();
            return lum > 0.62 ? Color.color(0.28, 0.28, 0.28) : Color.color(0.92, 0.92, 0.92);
        }
        double h = (background.getHue() + 180.0) % 360.0;
        double s = Math.min(1.0, Math.max(0.45, background.getSaturation() * 0.75 + 0.2));
        double lum =
                0.2126 * background.getRed()
                        + 0.7152 * background.getGreen()
                        + 0.0722 * background.getBlue();
        double bri = lum > 0.58 ? 0.28 : 0.9;
        return Color.hsb(h, s, bri);
    }

    private static Color resolveListBackground(Node styleAnchor) {
        if (styleAnchor == null) {
            return DEFAULT_LIST_BACKGROUND;
        }
        Scene scene = styleAnchor.getScene();
        if (scene != null) {
            Node popupList = styleAnchor.lookup(".combo-box-popup .list-view");
            if (popupList instanceof Region region
                    && region.getBackground() != null
                    && !region.getBackground().getFills().isEmpty()) {
                javafx.scene.paint.Paint paint = region.getBackground().getFills().getFirst().getFill();
                if (paint instanceof Color c) {
                    return c;
                }
            }
        }
        return DEFAULT_LIST_BACKGROUND;
    }

    private static String textFillStyle(Color fill) {
        return "-fx-font-size: 11px; -fx-fill: " + toCssHex(fill) + ";";
    }

    private static String toCssHex(Color color) {
        return String.format(
                Locale.ROOT,
                "#%02X%02X%02X",
                (int) Math.round(color.getRed() * 255),
                (int) Math.round(color.getGreen() * 255),
                (int) Math.round(color.getBlue() * 255));
    }

    static boolean[] highlightMask(String label, Collection<String> filterKeywords) {
        boolean[] mask = new boolean[label.length()];
        for (String keyword : nonEmptyKeywords(filterKeywords)) {
            markSubstringMatches(label, keyword.strip(), mask);
            String normalized = RequestFormMasterProductCandidateMatcher.normalize(keyword);
            if (!normalized.isEmpty() && !normalized.equals(keyword.strip().toUpperCase(Locale.ROOT))) {
                markSubstringMatches(label, normalized, mask);
            }
        }
        return mask;
    }

    private static List<String> nonEmptyKeywords(Collection<String> filterKeywords) {
        if (filterKeywords == null || filterKeywords.isEmpty()) {
            return List.of();
        }
        List<String> out = new ArrayList<>();
        for (String kw : filterKeywords) {
            if (kw != null && !kw.isBlank()) {
                out.add(kw);
            }
        }
        return out;
    }

    private static void markSubstringMatches(String label, String keyword, boolean[] mask) {
        if (keyword == null || keyword.isEmpty() || label.isEmpty()) {
            return;
        }
        String upperLabel = label.toUpperCase(Locale.ROOT);
        String upperKeyword = keyword.toUpperCase(Locale.ROOT);
        int from = 0;
        while (from <= upperLabel.length() - upperKeyword.length()) {
            int idx = upperLabel.indexOf(upperKeyword, from);
            if (idx < 0) {
                break;
            }
            for (int i = idx; i < idx + upperKeyword.length() && i < mask.length; i++) {
                mask[i] = true;
            }
            from = idx + 1;
        }
    }
}
