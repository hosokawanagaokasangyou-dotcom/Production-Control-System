package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
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

    @Test
    void exportExcelButtonExistsInToolBar() throws Exception {
        Element exportBtn = elementByFxId("exportExcelButton");
        assertNotNull(exportBtn, "exportExcelButton が無い");
        assertEquals("Excel出力", exportBtn.getAttribute("text"));
        assertEquals("#onExportExcelAction", exportBtn.getAttribute("onAction"));
    }

    @Test
    void openExcelButtonExistsInToolBar() throws Exception {
        Element openBtn = elementByFxId("openExcelButton");
        assertNotNull(openBtn, "openExcelButton が無い");
        assertEquals("Excelを開く", openBtn.getAttribute("text"));
        assertEquals("#onOpenExcelAction", openBtn.getAttribute("onAction"));
    }

    @Test
    void granularityTogglesExistInSubbar() throws Exception {
        Element daily = elementByFxId("granularityDailyToggle");
        Element monthly = elementByFxId("granularityMonthlyToggle");
        assertNotNull(daily, "granularityDailyToggle が無い");
        assertNotNull(monthly, "granularityMonthlyToggle が無い");
        assertEquals("日別", daily.getAttribute("text"));
        assertEquals("月別", monthly.getAttribute("text"));

        String css;
        try (InputStream in =
                ProcessingTrendTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/css/pm-ai-desktop.css")) {
            assertNotNull(in, "pm-ai-desktop.css が無い");
            css = new String(in.readAllBytes(), StandardCharsets.UTF_8);
        }
        assertTrue(
                css.contains(".pm-processing-trend-granularity-toggle .toggle-button"),
                "粒度トグルの CSS スタイルが無い");
    }

    @Test
    void actualSourceComboExistsInFilterBar() throws Exception {
        Element combo = elementByFxId("actualSourceCombo");
        assertNotNull(combo, "actualSourceCombo が無い");
    }

    @Test
    void comparisonColumnsExistInDetailTable() throws Exception {
        Element colCompare = elementByFxId("colCompareActual");
        assertNotNull(colCompare, "colCompareActual が無い");
        Element colDiff = elementByFxId("colActualCompareDiff");
        assertNotNull(colDiff, "colActualCompareDiff が無い");
        assertEquals("差異 (m)", colDiff.getAttribute("text"), "差異列の初期表示は『差異 (m)』であること");
        Element colCompareCum = elementByFxId("colCompareActualCum");
        assertNotNull(colCompareCum, "colCompareActualCum が無い");
        assertNull(elementByFxId("colDiff"), "予定との差異列 colDiff は削除されていること");
    }

    @Test
    void dailyLineChartAndMovingAverageColumnExist() throws Exception {
        Element lineChart = elementByFxId("dailyLineChart");
        assertNotNull(lineChart, "dailyLineChart が無い");
        Element colMa = elementByFxId("colActualMa");
        assertNotNull(colMa, "colActualMa が無い");
        assertEquals("30日移動平均 (m)", colMa.getAttribute("text"));
        assertNull(elementByFxId("colActual7dMa"), "旧 colActual7dMa は廃止されていること");
    }

    @Test
    void loadingChipIsLargeAndStyled() throws Exception {
        Element chip = elementByFxId("loadingChip");
        assertNotNull(chip, "loadingChip が無い");
        assertTrue(
                chip.getAttribute("styleClass").contains("pm-processing-trend-loading-chip"),
                chip.getAttribute("styleClass"));
        Element indicator = elementByFxId("loadingIndicator");
        assertNotNull(indicator, "loadingIndicator が無い");
        assertEquals("28.0", indicator.getAttribute("prefHeight"));
        assertEquals("28.0", indicator.getAttribute("prefWidth"));

        String css;
        try (InputStream in =
                ProcessingTrendTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/css/pm-ai-desktop.css")) {
            assertNotNull(in, "pm-ai-desktop.css が無い");
            css = new String(in.readAllBytes(), StandardCharsets.UTF_8);
        }
        String statusBlock = cssRuleBlock(css, ".label.pm-processing-trend-loading-status");
        Matcher size = Pattern.compile("-fx-font-size:\\s*(\\d+)px").matcher(statusBlock);
        assertTrue(size.find(), "読込ステータスの font-size が無い: " + statusBlock);
        int px = Integer.parseInt(size.group(1));
        assertTrue(px >= 14, "読込ステータスが小さすぎる: " + px + "px");
        assertTrue(css.contains(".pm-processing-trend-loading-chip"), "読込チップの CSS が無い");
    }

    @Test
    void movingAverageTogglesExistWith30DayDefault() throws Exception {
        Element ma7 = elementByFxId("ma7Toggle");
        Element ma14 = elementByFxId("ma14Toggle");
        Element ma30 = elementByFxId("ma30Toggle");
        assertNotNull(ma7, "ma7Toggle が無い");
        assertNotNull(ma14, "ma14Toggle が無い");
        assertNotNull(ma30, "ma30Toggle が無い");
        assertEquals("7日", ma7.getAttribute("text"));
        assertEquals("14日", ma14.getAttribute("text"));
        assertEquals("30日", ma30.getAttribute("text"));
        assertEquals("true", ma30.getAttribute("selected"), "既定は 30 日であること");
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
