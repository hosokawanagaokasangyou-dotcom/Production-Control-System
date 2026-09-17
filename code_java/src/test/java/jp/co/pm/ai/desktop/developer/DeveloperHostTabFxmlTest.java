package jp.co.pm.ai.desktop.developer;

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
import jp.co.pm.ai.desktop.config.MainShellTabLayoutDefaults;

class DeveloperHostTabFxmlTest {

    @Test
    void hostHasRuntimeErrorAndActionLogTabs() throws Exception {
        assertEquals(
                "jp.co.pm.ai.desktop.developer.DeveloperHostTabController", rootController());
        List<String> texts = tabTexts();
        assertTrue(texts.contains("実行時エラー"), texts.toString());
        assertTrue(texts.contains("操作ログ"), texts.toString());
        assertTrue(texts.contains("バックアップ"), texts.toString());
        assertNotNull(elementByFxId("runtimeErrorTab"));
        assertNotNull(elementByFxId("developerOperatorActionLogTab"));
        assertNotNull(elementByFxId("shareBackupTab"));
    }

    @Test
    void hostRootIsLayoutPaneNotTabPane() throws Exception {
        assertEquals("BorderPane", rootElementName());
        assertEquals("innerTabPane", elementByFxId("innerTabPane").getAttribute("fx:id"));
        assertEquals("TabPane", elementByFxId("innerTabPane").getTagName());
    }

    @Test
    void innerTabCatalogListsChildren() {
        assertEquals(
                List.of("実行時エラー", "操作ログ", "バックアップ"),
                MainShellInnerTabCatalog.labelsFor(MainShellTabId.DEVELOPER));
        assertTrue(MainShellInnerTabCatalog.titledPaneLabelsUnderInnerTab(MainShellTabId.DEVELOPER, 0).isEmpty());
        assertTrue(MainShellInnerTabCatalog.titledPaneLabelsUnderInnerTab(MainShellTabId.DEVELOPER, 1).isEmpty());
        assertTrue(MainShellInnerTabCatalog.titledPaneLabelsUnderInnerTab(MainShellTabId.DEVELOPER, 2).isEmpty());
    }

    @Test
    void shareBackupTabFxml_hasControllerAndPrimaryFxIds() throws Exception {
        try (InputStream in =
                DeveloperHostTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/DeveloperShareBackupTab.fxml")) {
            assertNotNull(in);
            var doc = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(in);
            Element root = doc.getDocumentElement();
            assertEquals(
                    "jp.co.pm.ai.desktop.developer.DeveloperShareBackupTabController",
                    root.getAttribute("fx:controller"));
        }
        assertNotNull(elementByFxIdInShareBackup("backupButton"));
        assertNotNull(elementByFxIdInShareBackup("openLocalButton"));
        assertNotNull(elementByFxIdInShareBackup("openKonanButton"));
        assertNotNull(elementByFxIdInShareBackup("openKokubuButton"));
        assertNotNull(elementByFxIdInShareBackup("loadingChip"));
        assertNotNull(elementByFxIdInShareBackup("konanSourceLabel"));
        assertNotNull(elementByFxIdInShareBackup("kokubuSharedLabel"));
        assertNotNull(elementByFxIdInShareBackup("statusLabel"));
        assertNotNull(elementByFxIdInShareBackup("logArea"));
    }

    @Test
    void mainShellTabTitleIsDeveloper() throws Exception {
        try (InputStream in =
                DeveloperHostTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/MainShell.fxml")) {
            assertNotNull(in);
            var doc = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(in);
            NodeList tabs = doc.getElementsByTagName("Tab");
            String title = null;
            for (int i = 0; i < tabs.getLength(); i++) {
                if (tabs.item(i) instanceof Element el
                        && "mainShellTabDeveloper".equals(el.getAttribute("fx:id"))) {
                    title = el.getAttribute("text");
                    break;
                }
            }
            assertEquals("開発", title);
        }
    }

    @Test
    void defaultLayoutIncludesDeveloperTab() {
        assertTrue(
                MainShellTabLayoutDefaults.DEFAULT_FLAT_TAB_KEY_ORDER.contains(
                        MainShellTabId.DEVELOPER.key()));
        List<String> leaves = new ArrayList<>();
        for (var n : MainShellTabLayoutDefaults.groupedLayout()) {
            collect(n, leaves);
        }
        assertTrue(leaves.contains(MainShellTabId.DEVELOPER.key()), leaves.toString());
    }

    private static void collect(
            jp.co.pm.ai.desktop.config.MainShellTabLayoutNode n, List<String> out) {
        if (n.isTab()) {
            out.add(n.id());
        } else {
            for (var c : n.children()) {
                collect(c, out);
            }
        }
    }

    private static String rootController() throws Exception {
        try (InputStream in =
                DeveloperHostTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/DeveloperHostTab.fxml")) {
            assertNotNull(in);
            var doc = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(in);
            return doc.getDocumentElement().getAttribute("fx:controller");
        }
    }

    private static String rootElementName() throws Exception {
        try (InputStream in =
                DeveloperHostTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/DeveloperHostTab.fxml")) {
            assertNotNull(in);
            var doc = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(in);
            return doc.getDocumentElement().getTagName();
        }
    }

    private static List<String> tabTexts() throws Exception {
        try (InputStream in =
                DeveloperHostTabFxmlTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/DeveloperHostTab.fxml")) {
            assertNotNull(in);
            var doc = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(in);
            NodeList tabs = doc.getElementsByTagName("Tab");
            List<String> out = new ArrayList<>();
            for (int i = 0; i < tabs.getLength(); i++) {
                if (tabs.item(i) instanceof Element el) {
                    String t = el.getAttribute("text");
                    if (t != null && !t.isBlank()) {
                        out.add(t);
                    }
                }
            }
            return out;
        }
    }

    private static Element elementByFxId(String fxId) throws Exception {
        return elementByFxIdIn(
                "/jp/co/pm/ai/desktop/fxml/DeveloperHostTab.fxml", fxId);
    }

    private static Element elementByFxIdInShareBackup(String fxId) throws Exception {
        return elementByFxIdIn(
                "/jp/co/pm/ai/desktop/fxml/DeveloperShareBackupTab.fxml", fxId);
    }

    private static Element elementByFxIdIn(String resource, String fxId) throws Exception {
        try (InputStream in = DeveloperHostTabFxmlTest.class.getResourceAsStream(resource)) {
            assertNotNull(in);
            var doc = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(in);
            NodeList all = doc.getElementsByTagName("*");
            for (int i = 0; i < all.getLength(); i++) {
                if (all.item(i) instanceof Element el && fxId.equals(el.getAttribute("fx:id"))) {
                    return el;
                }
            }
            return null;
        }
    }
}
