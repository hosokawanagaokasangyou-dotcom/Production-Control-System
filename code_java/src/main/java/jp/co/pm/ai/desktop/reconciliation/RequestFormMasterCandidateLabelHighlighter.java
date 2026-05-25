package jp.co.pm.ai.desktop.reconciliation;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Locale;

import javafx.geometry.Insets;
import javafx.scene.Node;
import javafx.scene.text.Text;
import javafx.scene.text.TextFlow;

/** マスタ候補コンボ表示文字列で、フィルタ語と一致する部分を黄色で強調する。 */
final class RequestFormMasterCandidateLabelHighlighter {

    private static final String PLAIN_STYLE = "-fx-font-size: 11px;";
    private static final String HIGHLIGHT_STYLE = "-fx-font-size: 11px; -fx-background-color: #FFEB3B;";

    private RequestFormMasterCandidateLabelHighlighter() {}

    static Node buildGraphic(String label, Collection<String> filterKeywords) {
        if (label == null || label.isEmpty()) {
            return new Text("");
        }
        List<String> keywords = nonEmptyKeywords(filterKeywords);
        if (keywords.isEmpty()) {
            Text plain = new Text(label);
            plain.setStyle(PLAIN_STYLE);
            return plain;
        }
        boolean[] highlight = highlightMask(label, keywords);
        TextFlow flow = new TextFlow();
        flow.setPadding(new Insets(1, 4, 1, 2));
        int i = 0;
        while (i < label.length()) {
            boolean on = highlight[i];
            int j = i + 1;
            while (j < label.length() && highlight[j] == on) {
                j++;
            }
            Text segment = new Text(label.substring(i, j));
            segment.setStyle(on ? HIGHLIGHT_STYLE : PLAIN_STYLE);
            flow.getChildren().add(segment);
            i = j;
        }
        return flow;
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
