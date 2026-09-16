package jp.co.pm.ai.desktop.kouchin;

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

import jp.co.pm.ai.desktop.MainShellTabId;
import jp.co.pm.ai.desktop.config.MainShellInnerTabCatalog;

class KouchinHostTabFxmlTest {

    @Test
    void hostHasVerifyTrendSourcesTabs() throws Exception {
        assertEquals(
                "jp.co.pm.ai.desktop.kouchin.KouchinHostTabController", rootController());
        List<String> texts = tabTexts();
        assertTrue(texts.contains("検証"), texts.toString());
        assertTrue(texts.contains("月次トレンド"), texts.toString());
        assertTrue(texts.contains("参照先"), texts.toString());
        assertNotNull(elementByFxId("verifyTab"));
        assertNotNull(elementByFxId("trendTab"));
        assertNotNull(elementByFxId("sourcesTab"));
    }

    @Test
    void innerTabCatalogListsChildren() {
        assertEquals(List.of("検証", "月次トレンド", "参照先"), MainShellInnerTabCatalog.labelsFor(MainShellTabId.KOUCHIN));
    }

    @Test
    void hostHasCenteredBusyOverlay() throws Exception {
        assertEquals("StackPane", rootTagName());
        assertNotNull(elementByFxId("busyOverlay"));
        assertNotNull(elementByFxId("busyIndicator"));
        assertNotNull(elementByFxId("busyBar"));
        assertNotNull(elementByFxId("busyLabel"));
        assertNotNull(elementByFxId("busyCancelButton"));
    }

    private static String rootController() throws Exception {
        try (InputStream in =
                KouchinHostTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/KouchinHostTab.fxml")) {
            assertNotNull(in);
            var doc = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(in);
            return doc.getDocumentElement().getAttribute("fx:controller");
        }
    }

    private static String rootTagName() throws Exception {
        try (InputStream in =
                KouchinHostTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/KouchinHostTab.fxml")) {
            assertNotNull(in);
            var doc = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(in);
            return doc.getDocumentElement().getTagName();
        }
    }

    private static Element elementByFxId(String id) throws Exception {
        try (InputStream in =
                KouchinHostTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/KouchinHostTab.fxml")) {
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
                KouchinHostTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/KouchinHostTab.fxml")) {
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
