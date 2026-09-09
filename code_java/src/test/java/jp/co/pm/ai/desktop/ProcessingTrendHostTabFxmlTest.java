package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import javax.xml.parsers.DocumentBuilderFactory;

import org.junit.jupiter.api.Test;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import jp.co.pm.ai.desktop.config.MainShellInnerTabCatalog;
import jp.co.pm.ai.desktop.MainShellTabId;

class ProcessingTrendHostTabFxmlTest {

    @Test
    void hostHasVolumeAndFeeChildTabs() throws Exception {
        assertEquals(
                "jp.co.pm.ai.desktop.ProcessingTrendHostTabController", rootController());
        List<String> texts = tabTexts();
        assertTrue(texts.contains("加工量"), texts.toString());
        assertTrue(texts.contains("加工賃"), texts.toString());
        assertNotNull(elementByFxId("volumeTab"));
        assertNotNull(elementByFxId("feeTab"));
    }

    @Test
    void innerTabCatalogListsChildren() {
        assertEquals(List.of("加工量", "加工賃"), MainShellInnerTabCatalog.labelsFor(MainShellTabId.PROCESSING_TREND));
    }

    private static String rootController() throws Exception {
        try (InputStream in =
                ProcessingTrendHostTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/ProcessingTrendHostTab.fxml")) {
            assertNotNull(in);
            var doc = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(in);
            return doc.getDocumentElement().getAttribute("fx:controller");
        }
    }

    private static Element elementByFxId(String id) throws Exception {
        try (InputStream in =
                ProcessingTrendHostTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/ProcessingTrendHostTab.fxml")) {
            assertNotNull(in);
            var doc = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(in);
            NodeList all = doc.getElementsByTagName("*");
            for (int i = 0; i < all.getLength(); i++) {
                if (all.item(i) instanceof Element el && id.equals(el.getAttribute("fx:id"))) {
                    return el;
                }
            }
            return null;
        }
    }

    private static List<String> tabTexts() throws Exception {
        try (InputStream in =
                ProcessingTrendHostTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/ProcessingTrendHostTab.fxml")) {
            assertNotNull(in);
            var doc = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(in);
            NodeList tabs = doc.getElementsByTagName("Tab");
            List<String> out = new ArrayList<>();
            for (int i = 0; i < tabs.getLength(); i++) {
                if (tabs.item(i) instanceof Element el) {
                    out.add(el.getAttribute("text"));
                }
            }
            return out;
        }
    }
}
