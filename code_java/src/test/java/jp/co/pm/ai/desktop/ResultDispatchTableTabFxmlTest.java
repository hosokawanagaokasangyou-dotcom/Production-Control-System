package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Map;
import java.util.stream.IntStream;

import javax.xml.parsers.DocumentBuilderFactory;

import org.junit.jupiter.api.Test;
import org.w3c.dom.Element;

class ResultDispatchTableTabFxmlTest {

    @Test
    void openLatestButtonHasDisabledBadgeCell() throws Exception {
        var resource =
                ResultDispatchTableTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/ResultDispatchTableTab.fxml");
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
        Map<String, String> buttonStyleById =
                buttonsById.entrySet().stream()
                        .collect(
                                java.util.stream.Collectors.toMap(
                                        Map.Entry::getKey,
                                        e -> e.getValue().getAttribute("styleClass")));

        var labels = document.getElementsByTagName("Label");
        java.util.Set<String> labelIds =
                IntStream.range(0, labels.getLength())
                        .mapToObj(labels::item)
                        .filter(Element.class::isInstance)
                        .map(Element.class::cast)
                        .filter(label -> label.hasAttribute("fx:id"))
                        .map(label -> label.getAttribute("fx:id"))
                        .collect(java.util.stream.Collectors.toSet());

        assertFalse(buttonStyleById.containsKey("aladdinEntryLocalExportButton"));
        assertFalse(buttonStyleById.containsKey("aladdinEntryOpenLocalLatestButton"));
        assertFalse(buttonStyleById.containsKey("aladdinEntryOpenLocalGenerationsButton"));
        assertFalse(buttonStyleById.containsKey("aladdinEntryExportButton"));
        assertTrue(buttonStyleById.containsKey("aladdinEntryReloadExportButton"));
        assertTrue(labelIds.contains("aladdinEntryReloadExportDisabledBadge"));
        assertTrue(buttonStyleById.containsKey("aladdinEntryOpenLatestButton"));
        assertTrue(labelIds.contains("aladdinEntryOpenLatestDisabledBadge"));
        assertTrue(buttonStyleById.containsKey("aladdinEntryIdentityCheckButton"));
        assertTrue(labelIds.contains("aladdinEntryIdentityCheckBadge"));
        assertFalse(
                "true".equalsIgnoreCase(
                        buttonsById.get("aladdinEntryIdentityCheckButton").getAttribute("disable")));
    }
}
