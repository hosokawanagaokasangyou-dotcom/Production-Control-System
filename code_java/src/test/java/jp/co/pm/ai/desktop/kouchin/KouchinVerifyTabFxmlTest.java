package jp.co.pm.ai.desktop.kouchin;

import static org.junit.jupiter.api.Assertions.assertEquals;
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
    @DisplayName("結果Excelが書けないときは書込不可バッジがある")
    void outputWriteBlockBadgeExists() throws Exception {
        Element badge = elementByFxId("outputWriteBlockBadge");
        assertNotNull(badge, "outputWriteBlockBadge が無い");
        assertTrue(
                badge.getAttribute("styleClass").contains("pm-kouchin-output-write-block-badge"),
                badge.getAttribute("styleClass"));
        assertEquals("Excel書込不可", badge.getAttribute("text"));
    }

    @Test
    @DisplayName("Excelを開くは国分と湖南でボタンが分かれている")
    void openExcelButtonsAreSplitByFactory() throws Exception {
        Element kokubu = elementByFxId("openKokubuExcelButton");
        Element konan = elementByFxId("openKonanExcelButton");
        assertNotNull(kokubu, "openKokubuExcelButton が無い");
        assertNotNull(konan, "openKonanExcelButton が無い");
        assertEquals("国分のExcelを開く", kokubu.getAttribute("text"));
        assertEquals("湖南のExcelを開く", konan.getAttribute("text"));
        assertEquals("#onOpenKokubuExcel", kokubu.getAttribute("onAction"));
        assertEquals("#onOpenKonanExcel", konan.getAttribute("onAction"));
    }

    @Test
    @DisplayName("検出表は国分と湖南の行が切れない高さ")
    void discoveryTableFitsFourRows() throws Exception {
        Element table = elementByFxId("discoveryTable");
        assertNotNull(table, "discoveryTable が無い");
        double pref = Double.parseDouble(table.getAttribute("prefHeight"));
        double min = table.getAttribute("minHeight").isBlank()
                ? pref
                : Double.parseDouble(table.getAttribute("minHeight"));
        assertTrue(pref >= 220.0, "prefHeight が両工場分より低い: " + pref);
        assertTrue(min >= 220.0, "minHeight が両工場分より低い: " + min);
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
