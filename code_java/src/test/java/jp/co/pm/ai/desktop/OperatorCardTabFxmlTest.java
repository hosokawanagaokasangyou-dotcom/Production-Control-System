package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Map;
import java.util.stream.IntStream;

import javax.xml.parsers.DocumentBuilderFactory;

import org.junit.jupiter.api.Test;
import org.w3c.dom.Element;

class OperatorCardTabFxmlTest {

    @Test
    void previewRefreshButtonPromptsForJsonSelection() throws Exception {
        var resource =
                OperatorCardTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/OperatorCardTab.fxml");
        var document = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(resource);
        var buttons = document.getElementsByTagName("Button");
        Map<String, Element> buttonsById =
                IntStream.range(0, buttons.getLength())
                        .mapToObj(buttons::item)
                        .filter(Element.class::isInstance)
                        .map(Element.class::cast)
                        .filter(button -> button.hasAttribute("fx:id"))
                        .collect(
                                java.util.stream.Collectors.toMap(
                                        button -> button.getAttribute("fx:id"),
                                        button -> button));

        assertTrue(buttonsById.containsKey("refreshPreviewButton"));
        assertEquals("プレビュー更新…", buttonsById.get("refreshPreviewButton").getAttribute("text"));
        assertEquals(
                "#onRefreshPreviewButtonAction",
                buttonsById.get("refreshPreviewButton").getAttribute("onAction"));
        assertTrue(buttonsById.containsKey("browseMemberButton"));
        assertTrue(buttonsById.containsKey("browseDispatchButton"));
    }
}
