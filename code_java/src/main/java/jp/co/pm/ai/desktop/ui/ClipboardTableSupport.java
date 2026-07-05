package jp.co.pm.ai.desktop.ui;

import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.StringReader;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import javafx.scene.input.ClipboardContent;

/** 表データを Excel / Outlook 等へ貼り付けやすい形式でクリップボードへ載せる。 */
public final class ClipboardTableSupport {

    private static volatile DataFlavor[] cachedHtmlTransferFlavors;

    private ClipboardTableSupport() {}

    /**
     * プレーンテキスト（TSV）と HTML 表の両方をクリップボードへ載せる。
     * Outlook 等のメール本文では HTML 側が表として貼り付く。
     */
    public static void copyTabularForRichTextPaste(String plainTsv, String tableHtml) {
        if (plainTsv == null || plainTsv.isBlank()) {
            return;
        }
        String htmlDoc = buildHtmlClipboardDocument(tableHtml != null ? tableHtml : "");
        String cfHtml = buildCfHtmlClipboardString(htmlDoc);
        if (isWindows()) {
            boolean awtOk = copyWindowsCfHtmlAndPlain(plainTsv, cfHtml);
            if (awtOk) {
                return;
            }
        }
        copyJavaFxHtmlAndPlain(htmlDoc, plainTsv);
    }

    public static String escapeHtml(String text) {
        if (text == null || text.isEmpty()) {
            return "";
        }
        StringBuilder sb = new StringBuilder(text.length() + 8);
        for (int i = 0; i < text.length(); i++) {
            char c = text.charAt(i);
            switch (c) {
                case '&' -> sb.append("&amp;");
                case '<' -> sb.append("&lt;");
                case '>' -> sb.append("&gt;");
                case '"' -> sb.append("&quot;");
                default -> sb.append(c);
            }
        }
        return sb.toString();
    }

    static String buildHtmlClipboardDocument(String tableHtml) {
        return "<html>\r\n"
                + "<head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\"></head>\r\n"
                + "<body>\r\n"
                + "<!--StartFragment-->\r\n"
                + tableHtml
                + "\r\n<!--EndFragment-->\r\n"
                + "</body>\r\n"
                + "</html>";
    }

    /**
     * Windows Outlook / Word 向け CF_HTML 形式。
     * オフセット付きヘッダを明示する（{@code StartHTML} 等）。
     */
    static String buildCfHtmlClipboardString(String htmlDocument) {
        String body = htmlDocument != null ? htmlDocument : "";
        String headerPlaceholder =
                "Version:1.0\r\n"
                        + "StartHTML:0000000000\r\n"
                        + "EndHTML:0000000000\r\n"
                        + "StartFragment:0000000000\r\n"
                        + "EndFragment:0000000000\r\n";
        byte[] draftBytes = (headerPlaceholder + body).getBytes(StandardCharsets.UTF_8);
        int startHtml = headerPlaceholder.getBytes(StandardCharsets.UTF_8).length;
        int endHtml = draftBytes.length;
        int startFragment = startHtml;
        int endFragment = endHtml;
        byte[] startMarker = "<!--StartFragment-->".getBytes(StandardCharsets.UTF_8);
        byte[] endMarker = "<!--EndFragment-->".getBytes(StandardCharsets.UTF_8);
        int fragStart = indexOfUtf8(draftBytes, startMarker);
        if (fragStart >= 0) {
            startFragment = fragStart + startMarker.length;
            while (startFragment < endHtml
                    && (draftBytes[startFragment] == '\r' || draftBytes[startFragment] == '\n')) {
                startFragment++;
            }
        }
        int fragEnd = indexOfUtf8(draftBytes, endMarker);
        if (fragEnd >= 0) {
            endFragment = fragEnd;
        }
        String header =
                String.format(
                        Locale.ROOT,
                        "Version:1.0\r\nStartHTML:%010d\r\nEndHTML:%010d\r\nStartFragment:%010d\r\nEndFragment:%010d\r\n",
                        startHtml,
                        endHtml,
                        startFragment,
                        endFragment);
        return header + body;
    }

    private static boolean copyWindowsCfHtmlAndPlain(String plainTsv, String cfHtml) {
        try {
            Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
            clipboard.setContents(new HtmlAndPlainTransferable(plainTsv, cfHtml), null);
            return true;
        } catch (Exception ex) {
            return false;
        }
    }

    private static void copyJavaFxHtmlAndPlain(String htmlDoc, String plainTsv) {
        ClipboardContent content = new ClipboardContent();
        content.putString(plainTsv);
        content.putHtml(htmlDoc);
        javafx.scene.input.Clipboard.getSystemClipboard().setContent(content);
    }

    static DataFlavor[] htmlTransferFlavors() {
        DataFlavor[] cached = cachedHtmlTransferFlavors;
        if (cached != null) {
            return cached;
        }
        List<DataFlavor> flavors = new ArrayList<>();
        flavors.add(DataFlavor.stringFlavor);
        addFlavorIfValid(flavors, "HTML Format;class=\"[B\"");
        addFlavorIfValid(flavors, "text/html; charset=UTF-8; class=java.lang.String");
        addFlavorIfValid(flavors, "text/html; charset=UTF-8; class=java.io.InputStream");
        addFlavorIfValid(flavors, "text/html; class=java.lang.String");
        try {
            Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
            for (DataFlavor flavor : clipboard.getAvailableDataFlavors()) {
                if (isRegisteredHtmlFormatName(flavor) && !flavors.contains(flavor)) {
                    flavors.add(flavor);
                }
            }
        } catch (Exception ignored) {
            // headless / toolkit unavailable
        }
        cached = flavors.toArray(DataFlavor[]::new);
        cachedHtmlTransferFlavors = cached;
        return cached;
    }

    private static void addFlavorIfValid(List<DataFlavor> flavors, String mime) {
        try {
            flavors.add(new DataFlavor(mime));
        } catch (Exception ignored) {
            // invalid mime for this JVM — skip
        }
    }

    private static boolean isRegisteredHtmlFormatName(DataFlavor flavor) {
        if (flavor == null) {
            return false;
        }
        String name = flavor.getHumanPresentableName();
        if (name != null && name.toLowerCase(Locale.ROOT).contains("html format")) {
            return true;
        }
        String mime = flavor.getMimeType();
        return mime != null && mime.toLowerCase(Locale.ROOT).contains("html");
    }

    static boolean isHtmlClipboardFlavor(DataFlavor flavor) {
        if (flavor == null || DataFlavor.stringFlavor.equals(flavor)) {
            return false;
        }
        for (DataFlavor registered : htmlTransferFlavors()) {
            if (registered.equals(flavor)) {
                return !DataFlavor.stringFlavor.equals(flavor);
            }
        }
        return isRegisteredHtmlFormatName(flavor);
    }

    static Object htmlTransferData(DataFlavor flavor, String cfHtml)
            throws UnsupportedFlavorException, IOException {
        if (DataFlavor.stringFlavor.equals(flavor)) {
            throw new UnsupportedFlavorException(flavor);
        }
        if (!isHtmlClipboardFlavor(flavor)) {
            throw new UnsupportedFlavorException(flavor);
        }
        Class<?> rep = flavor.getRepresentationClass();
        if (rep != null && byte[].class.equals(rep)) {
            return cfHtml.getBytes(StandardCharsets.UTF_8);
        }
        if (rep != null && InputStream.class.isAssignableFrom(rep)) {
            return new ByteArrayInputStream(cfHtml.getBytes(StandardCharsets.UTF_8));
        }
        if (rep != null
                && (StringReader.class.isAssignableFrom(rep)
                        || java.io.Reader.class.isAssignableFrom(rep))) {
            return new StringReader(cfHtml);
        }
        return cfHtml;
    }

    private static int indexOfUtf8(byte[] haystack, byte[] needle) {
        if (needle.length == 0 || haystack.length < needle.length) {
            return -1;
        }
        outer:
        for (int i = 0; i <= haystack.length - needle.length; i++) {
            for (int j = 0; j < needle.length; j++) {
                if (haystack[i + j] != needle[j]) {
                    continue outer;
                }
            }
            return i;
        }
        return -1;
    }

    private static boolean isWindows() {
        String os = System.getProperty("os.name", "");
        return os.toLowerCase(Locale.ROOT).contains("win");
    }

    private static final class HtmlAndPlainTransferable implements Transferable {
        private final String plain;
        private final String cfHtml;
        private final DataFlavor[] flavors;

        private HtmlAndPlainTransferable(String plain, String cfHtml) {
            this.plain = plain != null ? plain : "";
            this.cfHtml = cfHtml != null ? cfHtml : "";
            this.flavors = htmlTransferFlavors();
        }

        @Override
        public DataFlavor[] getTransferDataFlavors() {
            return flavors.clone();
        }

        @Override
        public boolean isDataFlavorSupported(DataFlavor flavor) {
            if (DataFlavor.stringFlavor.equals(flavor)) {
                return true;
            }
            return isHtmlClipboardFlavor(flavor);
        }

        @Override
        public Object getTransferData(DataFlavor flavor)
                throws UnsupportedFlavorException, IOException {
            if (DataFlavor.stringFlavor.equals(flavor)) {
                return plain;
            }
            return htmlTransferData(flavor, cfHtml);
        }
    }
}
