package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.awt.datatransfer.DataFlavor;
import java.nio.charset.StandardCharsets;

import org.junit.jupiter.api.Test;

class ClipboardTableSupportTest {

    @Test
    void buildCfHtmlClipboardString_includesValidOffsets() {
        String table = "<table><tr><td>A</td></tr></table>";
        String doc = ClipboardTableSupport.buildHtmlClipboardDocument(table);
        String cf = ClipboardTableSupport.buildCfHtmlClipboardString(doc);
        assertTrue(cf.startsWith("Version:1.0\r\n"));
        assertTrue(cf.contains("StartHTML:"));
        assertTrue(cf.contains("<!--StartFragment-->"));
        assertTrue(cf.contains(table));
        assertOffsetConsistency(cf);
    }

    @Test
    void buildCfHtmlClipboardString_supportsJapaneseContent() {
        String table = "<table><tr><td>スリット機1　湖南</td></tr></table>";
        String doc = ClipboardTableSupport.buildHtmlClipboardDocument(table);
        String cf = ClipboardTableSupport.buildCfHtmlClipboardString(doc);
        assertTrue(cf.contains("スリット機1　湖南"));
        assertOffsetConsistency(cf);
    }

    @Test
    void escapeHtml_escapesSpecialChars() {
        assertEquals("a&amp;b&lt;c", ClipboardTableSupport.escapeHtml("a&b<c"));
    }

    @Test
    void buildHtmlClipboardDocument_wrapsFragment() {
        String doc = ClipboardTableSupport.buildHtmlClipboardDocument("<table></table>");
        assertTrue(doc.contains("<!--StartFragment-->"));
        assertTrue(doc.contains("<table></table>"));
        assertTrue(doc.contains("<!--EndFragment-->"));
    }

    @Test
    void htmlTransferFlavors_doesNotThrowOnClassInit() {
        DataFlavor[] flavors = ClipboardTableSupport.htmlTransferFlavors();
        assertTrue(flavors.length >= 1);
        assertEquals(DataFlavor.stringFlavor, flavors[0]);
    }

    @Test
    void htmlTransferData_returnsStringOrBytesForHtmlFlavor() throws Exception {
        String cf = ClipboardTableSupport.buildCfHtmlClipboardString(
                ClipboardTableSupport.buildHtmlClipboardDocument("<table></table>"));
        for (DataFlavor flavor : ClipboardTableSupport.htmlTransferFlavors()) {
            if (DataFlavor.stringFlavor.equals(flavor)) {
                continue;
            }
            Object data = ClipboardTableSupport.htmlTransferData(flavor, cf);
            assertTrue(
                    data instanceof String
                            || data instanceof byte[]
                            || data instanceof java.io.InputStream
                            || data instanceof java.io.Reader);
            if (data instanceof byte[] bytes) {
                assertTrue(new String(bytes, StandardCharsets.UTF_8).contains("<table"));
            }
            if (data instanceof String text) {
                assertTrue(text.contains("<table"));
            }
        }
    }

    private static void assertOffsetConsistency(String cfHtml) {
        byte[] bytes = cfHtml.getBytes(StandardCharsets.UTF_8);
        int startHtml = parseOffset(cfHtml, "StartHTML:");
        int endHtml = parseOffset(cfHtml, "EndHTML:");
        int startFragment = parseOffset(cfHtml, "StartFragment:");
        int endFragment = parseOffset(cfHtml, "EndFragment:");
        assertTrue(startHtml >= 0 && startHtml < bytes.length);
        assertEquals(bytes.length, endHtml);
        assertTrue(startFragment >= startHtml && startFragment <= endHtml);
        assertTrue(endFragment >= startFragment && endFragment <= endHtml);
        assertEquals('<', (char) bytes[startHtml]);
        String fragment =
                new String(bytes, startFragment, endFragment - startFragment, StandardCharsets.UTF_8);
        assertTrue(fragment.contains("<table"));
    }

    private static int parseOffset(String cfHtml, String key) {
        int keyIdx = cfHtml.indexOf(key);
        if (keyIdx < 0) {
            return -1;
        }
        int lineEnd = cfHtml.indexOf('\n', keyIdx);
        String line = cfHtml.substring(keyIdx + key.length(), lineEnd >= 0 ? lineEnd : cfHtml.length())
                .strip();
        return Integer.parseInt(line);
    }
}
