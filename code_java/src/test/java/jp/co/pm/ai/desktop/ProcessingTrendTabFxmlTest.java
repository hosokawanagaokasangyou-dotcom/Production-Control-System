package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.parsers.DocumentBuilderFactory;

import org.junit.jupiter.api.Test;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

class ProcessingTrendTabFxmlTest {

    @Test
    void sourceSummaryLabelIsImmediatelyBelowChartStack() throws Exception {
        Element source = elementByFxId("sourceSummaryLabel");
        assertNotNull(source, "sourceSummaryLabel が無い");
        Element prev = previousElementSibling(source);
        assertNotNull(prev, "sourceSummaryLabel の直前に兄弟要素が無い");
        assertEquals("chartStack", prev.getAttribute("fx:id"), "ソース表示はグラフ直下であるべき");
    }

    @Test
    void legendBoxHasDedicatedStyleClass() throws Exception {
        Element legend = elementByFxId("legendBox");
        assertNotNull(legend, "legendBox が無い");
        String styleClass = legend.getAttribute("styleClass");
        assertTrue(styleClass.contains("pm-processing-trend-legend"), styleClass);
    }

    @Test
    void legendCssIsLargeOnWhiteBackground() throws Exception {
        String css;
        try (InputStream in =
                ProcessingTrendTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/css/pm-ai-desktop.css")) {
            assertNotNull(in, "pm-ai-desktop.css が無い");
            css = new String(in.readAllBytes(), StandardCharsets.UTF_8);
        }

        String legendBlock = cssRuleBlock(css, ".pm-processing-trend-legend");
        assertTrue(
                legendBlock.contains("-fx-background-color: white")
                        || legendBlock.contains("-fx-background-color: -pm-trend-chart-bg"),
                "凡例背景が白ではない: " + legendBlock);

        String labelBlock = cssRuleBlock(css, ".label.pm-legend-label");
        Matcher size = Pattern.compile("-fx-font-size:\\s*(\\d+)px").matcher(labelBlock);
        assertTrue(size.find(), "凡例の font-size が無い: " + labelBlock);
        int px = Integer.parseInt(size.group(1));
        assertTrue(px >= 14, "凡例が小さすぎる: " + px + "px");
        assertTrue(
                labelBlock.contains("-fx-text-fill:"),
                "白背景に対する凡例文字色が未指定: " + labelBlock);
    }

    private static Element elementByFxId(String fxId) throws Exception {
        try (InputStream in =
                ProcessingTrendTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/ProcessingTrendTab.fxml")) {
            var document = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(in);
            NodeList all = document.getElementsByTagName("*");
            for (int i = 0; i < all.getLength(); i++) {
                if (!(all.item(i) instanceof Element el)) {
                    continue;
                }
                if (fxId.equals(el.getAttribute("fx:id"))) {
                    return el;
                }
            }
            return null;
        }
    }

    private static Element previousElementSibling(Element el) {
        Node n = el.getPreviousSibling();
        while (n != null) {
            if (n instanceof Element e) {
                return e;
            }
            n = n.getPreviousSibling();
        }
        return null;
    }

    private static String cssRuleBlock(String css, String selector) {
        int idx = css.indexOf(selector);
        assertTrue(idx >= 0, "CSS セレクタが無い: " + selector);
        int brace = css.indexOf('{', idx);
        int end = css.indexOf('}', brace);
        assertTrue(brace >= 0 && end > brace, "CSS ルールが閉じられていない: " + selector);
        return css.substring(brace, end + 1);
    }
}
