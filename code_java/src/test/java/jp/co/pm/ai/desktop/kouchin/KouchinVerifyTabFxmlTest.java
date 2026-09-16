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
        assertTrue(tooltipText(kokubu).contains("読み取り専用"), tooltipText(kokubu));
        assertTrue(tooltipText(konan).contains("読み取り専用"), tooltipText(konan));
    }

    @Test
    @DisplayName("対象月の②を読み書き可で開くボタンがある")
    void openSource2ButtonsAreSplitByFactory() throws Exception {
        Element kokubu = elementByFxId("openKokubuSource2Button");
        Element konan = elementByFxId("openKonanSource2Button");
        assertNotNull(kokubu, "openKokubuSource2Button が無い");
        assertNotNull(konan, "openKonanSource2Button が無い");
        assertEquals("国分の②を開く", kokubu.getAttribute("text"));
        assertEquals("湖南の②を開く", konan.getAttribute("text"));
        assertEquals("#onOpenKokubuSource2", kokubu.getAttribute("onAction"));
        assertEquals("#onOpenKonanSource2", konan.getAttribute("onAction"));
        assertTrue(tooltipText(kokubu).contains("読み書き可"), tooltipText(kokubu));
        assertTrue(tooltipText(konan).contains("読み書き可"), tooltipText(konan));
    }

    @Test
    @DisplayName("ツールバーは折り返し、判定フィルタは結果表の直前")
    void toolbarWrapsAndFilterSitsAboveResults() throws Exception {
        try (InputStream in = KouchinVerifyTabFxmlTest.class.getResourceAsStream(
                "/jp/co/pm/ai/desktop/fxml/KouchinVerifyTab.fxml")) {
            assertNotNull(in);
            var doc = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(in);
            boolean hasFlow = false;
            NodeList all = doc.getElementsByTagName("*");
            int discovery = -1;
            int search = -1;
            int result = -1;
            for (int i = 0; i < all.getLength(); i++) {
                if (!(all.item(i) instanceof Element el)) {
                    continue;
                }
                if ("FlowPane".equals(el.getTagName()) && el.getElementsByTagName("Button").getLength() > 0) {
                    hasFlow = true;
                }
                String id = el.getAttribute("fx:id");
                if ("discoveryTable".equals(id)) {
                    discovery = i;
                } else if ("searchField".equals(id)) {
                    search = i;
                } else if ("resultTable".equals(id)) {
                    result = i;
                }
            }
            assertTrue(hasFlow, "ツールバーが FlowPane ではない");
            assertTrue(discovery >= 0 && search > discovery && result > search,
                    "判定フィルタが結果表の直前にない discovery=" + discovery + " search=" + search + " result=" + result);
        }
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

    @Test
    @DisplayName("対象月の自動選択理由を出すラベルがある")
    void targetYmReasonLabelExists() throws Exception {
        Element label = elementByFxId("targetYmReasonLabel");
        assertNotNull(label, "targetYmReasonLabel が無い");
        assertTrue(label.getAttribute("wrapText").contains("true") || "true".equals(label.getAttribute("wrapText")));
        assertTrue(label.getAttribute("text").contains("RVSHEET") || label.getAttribute("text").contains("対象月"),
                label.getAttribute("text"));
    }

    @Test
    @DisplayName("取込ボタンは東レ提供CSVと分かる")
    void importCsvButtonNamesTorayCsv() throws Exception {
        Element button = elementByFxId("importCsvButton");
        assertNotNull(button, "importCsvButton が無い");
        assertEquals("東レ提供CSVを取り込む…", button.getAttribute("text"));
    }

    @Test
    @DisplayName("③の案内は月次実績とする")
    void source3HintSaysMonthlyActual() throws Exception {
        Element hint = elementByFxId("diffHintLabel");
        assertNotNull(hint, "diffHintLabel が無い");
        assertTrue(hint.getAttribute("text").contains("③月次実績"), hint.getAttribute("text"));
        assertTrue(!hint.getAttribute("text").contains("③アラジン"), hint.getAttribute("text"));
    }

    private static String tooltipText(Element button) {
        NodeList tips = button.getElementsByTagName("Tooltip");
        if (tips.getLength() == 0) {
            return "";
        }
        return ((Element) tips.item(0)).getAttribute("text");
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
