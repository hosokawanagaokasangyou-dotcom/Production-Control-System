package jp.co.pm.ai.desktop.reconciliation;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Locale;

import javafx.geometry.Insets;
import javafx.scene.Node;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;
import javafx.scene.text.FontWeight;
import javafx.scene.text.Text;
import javafx.scene.text.TextFlow;

/**
 * マスタ候補コンボ表示文字列の着色。
 * <ul>
 *   <li>フィルタ不一致 … フォーム見出し（.label）と同系の明色
 *   <li>フィルタ一致 … 黄色（#FFEB3B）
 * </ul>
 */
final class RequestFormMasterCandidateLabelHighlighter {

    /** フォーム見出し .label と同系（midnight-blue の -fx-dark-text-color 相当）。 */
    static final String LABEL_TEXT_FILL_HEX = "#EFF6FF";

    /** フィルタ一致部分の文字色（黄色）。 */
    static final String HIGHLIGHT_TEXT_FILL_HEX = "#FFEB3B";

    private static final Color PLAIN_FILL = Color.web(LABEL_TEXT_FILL_HEX);
    private static final Color HIGHLIGHT_FILL = Color.web(HIGHLIGHT_TEXT_FILL_HEX);
    private static final Font CANDIDATE_FONT = Font.font(11);

    private RequestFormMasterCandidateLabelHighlighter() {}

    static Node buildGraphic(String label, Collection<String> filterKeywords) {
        if (label == null || label.isEmpty()) {
            return new Text("");
        }
        List<String> keywords = nonEmptyKeywords(filterKeywords);
        if (keywords.isEmpty()) {
            return plainText(label);
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
            flow.getChildren().add(styledText(label.substring(i, j), on));
            i = j;
        }
        return flow;
    }

    private static Text plainText(String text) {
        Text node = new Text(text);
        node.setFont(CANDIDATE_FONT);
        node.setFill(PLAIN_FILL);
        node.getStyleClass().add("request-form-master-candidate-plain");
        return node;
    }

    private static Text styledText(String text, boolean highlighted) {
        Text node = new Text(text);
        node.setFont(highlighted ? Font.font(CANDIDATE_FONT.getFamily(), FontWeight.BOLD, 11) : CANDIDATE_FONT);
        node.setFill(highlighted ? HIGHLIGHT_FILL : PLAIN_FILL);
        node.getStyleClass().add(
                highlighted ? "request-form-master-candidate-highlight" : "request-form-master-candidate-plain");
        return node;
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
