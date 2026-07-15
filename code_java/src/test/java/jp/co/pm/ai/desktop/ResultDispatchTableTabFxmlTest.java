package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Map;
import java.util.stream.IntStream;

import javax.xml.parsers.DocumentBuilderFactory;

import org.junit.jupiter.api.Test;
import org.w3c.dom.Element;

class ResultDispatchTableTabFxmlTest {

    @Test
    void localOpenButtonsUseTheSameSubduedStyleAsLocalExport() throws Exception {
        var resource =
                ResultDispatchTableTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/ResultDispatchTableTab.fxml");
        var document = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(resource);
        var buttons = document.getElementsByTagName("Button");
        Map<String, String> styleById =
                IntStream.range(0, buttons.getLength())
                        .mapToObj(buttons::item)
                        .filter(Element.class::isInstance)
                        .map(Element.class::cast)
                        .filter(button -> button.hasAttribute("fx:id"))
                        .collect(
                                java.util.stream.Collectors.toMap(
                                        button -> button.getAttribute("fx:id"),
                                        button -> button.getAttribute("styleClass")));

        String localExportStyle = styleById.get("aladdinEntryLocalExportButton");
        assertEquals(localExportStyle, styleById.get("aladdinEntryOpenLocalLatestButton"));
        assertEquals(localExportStyle, styleById.get("aladdinEntryOpenLocalGenerationsButton"));
        assertFalse(styleById.containsKey("aladdinEntryExportButton"));
        assertTrue(styleById.containsKey("aladdinEntryReloadExportButton"));
    }
}
