package jp.co.pm.ai.desktop.kouchin;

import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.InputStream;

import javax.xml.parsers.DocumentBuilderFactory;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

class KouchinVerifyTabFxmlTest {

    @Test
    @DisplayName("検出表は①②③＋湖南月次の4行がヘッダごと見切れない高さ")
    void discoveryTableFitsFourRows() throws Exception {
        Element table = elementByFxId("discoveryTable");
        assertNotNull(table, "discoveryTable が無い");
        double pref = Double.parseDouble(table.getAttribute("prefHeight"));
        double min = table.getAttribute("minHeight").isBlank()
                ? pref
                : Double.parseDouble(table.getAttribute("minHeight"));
        assertTrue(pref >= 148.0, "prefHeight が4行分より低い: " + pref);
        assertTrue(min >= 148.0, "minHeight が4行分より低い: " + min);
    }

    private static Element elementByFxId(String id) throws Exception {
        try (InputStream in =
                KouchinVerifyTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/KouchinVerifyTab.fxml")) {
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
}
