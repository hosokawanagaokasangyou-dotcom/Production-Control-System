package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;

import java.io.InputStream;
import java.util.List;

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
        assertNull(elementByFxId("colDiffYen"), "差異列は表示しない");
        assertNull(elementByFxId("colReqDiffYen"), "差異列は表示しない");
        assertNotNull(elementByFxId("colActualCumYen"));
        assertNotNull(elementByFxId("colPlanCumYen"));
        assertNotNull(elementByFxId("requestPane"));
        assertNotNull(elementByFxId("planSourceCombo"));
        assertNotNull(elementByFxId("planSourceMetaLabel"));
        assertNotNull(elementByFxId("sourceSummaryLabel"));
        assertNotNull(elementByFxId("colRequestNo"));
        assertNotNull(elementByFxId("colAoYen"));
        assertNotNull(elementByFxId("colRateYen"));
        assertNotNull(elementByFxId("colReqActualYen"));
        assertNotNull(elementByFxId("kpiOrderAoYen"));
        assertNotNull(elementByFxId("prevPeriodButton"));
        assertNotNull(elementByFxId("nextPeriodButton"));
        assertNotNull(elementByFxId("exportExcelButton"));
        assertNotNull(elementByFxId("openExcelButton"));
        assertNotNull(elementByFxId("autoRefreshCheckBox"));
        assertNotNull(elementByFxId("autoRefreshIntervalSpinner"));
        assertNotNull(elementByFxId("nextRefreshLabel"));
        assertNotNull(elementByFxId("viewComboToggle"));
        assertNotNull(elementByFxId("viewDailyToggle"));
        assertNotNull(elementByFxId("viewCumulativeToggle"));
        assertEquals("AO(受注額)", elementByFxId("colAoYen").getAttribute("text"));
        assertEquals("円/m", elementByFxId("colRateYen").getAttribute("text"));
        assertEquals("未了円", elementByFxId("colPlanYen").getAttribute("text"));
        assertEquals("未了円", elementByFxId("colReqPlanYen").getAttribute("text"));
        assertEquals("未了 (m)", elementByFxId("colPlanM").getAttribute("text"));
        assertEquals("false", elementByFxId("detailPane").getAttribute("expanded"));
        assertEquals("false", elementByFxId("requestPane").getAttribute("expanded"));
        assertEquals(
                "jp.co.pm.ai.desktop.ProcessingFeeTrendTabController",
                rootController());
        assertFxmlImportsControls();
    }

    @Test
    void fxmlImportsTitledPaneAndTableTypes() throws Exception {
        assertFxmlImportsControls();
    }

    /**
     * Chart 系（BarChart / LineChart）は軸を @NamedArg で受けるため FXMLLoader が ProxyBuilder 経由で生成し、
     * その経路では styleClass のカンマ区切りが分割されず「a, b」が 1 つのクラス名になる（CSS が一切当たらない）。
     */
    @Test
    void chartStyleClassMustBeSingleToken() throws Exception {
        try (InputStream in =
                ProcessingFeeTrendTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/ProcessingFeeTrendTab.fxml")) {
            assertNotNull(in);
            var doc = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(in);
            for (String tag : List.of("BarChart", "LineChart")) {
                NodeList charts = doc.getElementsByTagName(tag);
                for (int i = 0; i < charts.getLength(); i++) {
                    Element el = (Element) charts.item(i);
                    String sc = el.getAttribute("styleClass");
                    assertFalse(
                            sc.contains(",") || sc.trim().contains(" "),
                            tag + " styleClass はカンマ区切り不可（ProxyBuilder は分割しない）: " + sc);
                }
            }
        }
    }

    private static void assertFxmlImportsControls() throws Exception {
        try (InputStream in =
                ProcessingFeeTrendTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/ProcessingFeeTrendTab.fxml")) {
            assertNotNull(in);
            String xml = new String(in.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8);
            org.junit.jupiter.api.Assertions.assertTrue(
                    xml.contains("<?import javafx.scene.control.CheckBox?>"),
                    "CheckBox import required for auto-refresh");
            org.junit.jupiter.api.Assertions.assertTrue(
                    xml.contains("<?import javafx.scene.control.Spinner?>"),
                    "Spinner import required for auto-refresh interval");
            org.junit.jupiter.api.Assertions.assertTrue(
                    xml.contains("<?import javafx.scene.control.ToggleButton?>"),
                    "ToggleButton import required for view mode");
            org.junit.jupiter.api.Assertions.assertTrue(
                    xml.contains("<?import javafx.scene.control.TitledPane?>"),
                    "TitledPane import required for FXMLLoader");
            org.junit.jupiter.api.Assertions.assertFalse(
                    xml.contains("<?import javafx.scene.control.ScrollPane?>"),
                    "ScrollPane breaks chart white background inheritance; keep volume-like VBox center");
            org.junit.jupiter.api.Assertions.assertFalse(
                    xml.contains("<ScrollPane"),
                    "ScrollPane element must not wrap chart center");
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
