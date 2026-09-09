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
        assertNotNull(elementByFxId("detailPane"));
        assertNotNull(elementByFxId("detailTable"));
        assertNotNull(elementByFxId("colActualYen"));
        assertNotNull(elementByFxId("colPlanYen"));
        assertNotNull(elementByFxId("colDiffYen"));
        assertNotNull(elementByFxId("colActualCumYen"));
        assertNotNull(elementByFxId("colPlanCumYen"));
        assertNotNull(elementByFxId("requestPane"));
        assertNotNull(elementByFxId("requestTable"));
        assertNotNull(elementByFxId("colRequestNo"));
        assertNotNull(elementByFxId("colRateYen"));
        assertNotNull(elementByFxId("colReqActualYen"));
        assertEquals(
                "jp.co.pm.ai.desktop.ProcessingFeeTrendTabController",
                rootController());
        assertFxmlImportsControls();
    }

    @Test
    void fxmlImportsTitledPaneAndTableTypes() throws Exception {
        assertFxmlImportsControls();
    }

    private static void assertFxmlImportsControls() throws Exception {
        try (InputStream in =
                ProcessingFeeTrendTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/ProcessingFeeTrendTab.fxml")) {
            assertNotNull(in);
            String xml = new String(in.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8);
            org.junit.jupiter.api.Assertions.assertTrue(
                    xml.contains("<?import javafx.scene.control.TitledPane?>"),
                    "TitledPane import required for FXMLLoader");
            org.junit.jupiter.api.Assertions.assertTrue(
                    xml.contains("<?import javafx.scene.control.TableView?>"),
                    "TableView import required for FXMLLoader");
            org.junit.jupiter.api.Assertions.assertTrue(
                    xml.contains("<?import javafx.scene.control.TableColumn?>"),
                    "TableColumn import required for FXMLLoader");
        }
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
