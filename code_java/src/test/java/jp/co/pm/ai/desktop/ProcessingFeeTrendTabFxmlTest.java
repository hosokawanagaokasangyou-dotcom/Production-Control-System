package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

import java.io.InputStream;

import javax.xml.parsers.DocumentBuilderFactory;

import org.junit.jupiter.api.Test;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

class ProcessingFeeTrendTabFxmlTest {

    @Test
    void fxmlHasChartStackAndKpis() throws Exception {
        assertNotNull(elementByFxId("chartStack"));
        assertNotNull(elementByFxId("kpiActualYen"));
        assertNotNull(elementByFxId("kpiPlanYen"));
        assertNotNull(elementByFxId("dailyChart"));
        assertNotNull(elementByFxId("cumulativeChart"));
        assertNotNull(elementByFxId("dailyYAxis"));
        assertNotNull(elementByFxId("cumulativeYAxis"));
        assertEquals(
                "jp.co.pm.ai.desktop.ProcessingFeeTrendTabController",
                rootController());
    }

    private static String rootController() throws Exception {
        try (InputStream in =
                ProcessingFeeTrendTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/ProcessingFeeTrendTab.fxml")) {
            assertNotNull(in);
            var doc =
                    DocumentBuilderFactory.newInstance()
                            .newDocumentBuilder()
                            .parse(in);
            return doc.getDocumentElement().getAttribute("fx:controller");
        }
    }

    private static Element elementByFxId(String id) throws Exception {
        try (InputStream in =
                ProcessingFeeTrendTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/ProcessingFeeTrendTab.fxml")) {
            assertNotNull(in);
            var doc =
                    DocumentBuilderFactory.newInstance()
                            .newDocumentBuilder()
                            .parse(in);
            NodeList all = doc.getElementsByTagName("*");
            for (int i = 0; i < all.getLength(); i++) {
                if (all.item(i) instanceof Element el && id.equals(el.getAttribute("fx:id"))) {
                    return el;
                }
            }
            return null;
        }
    }
}
