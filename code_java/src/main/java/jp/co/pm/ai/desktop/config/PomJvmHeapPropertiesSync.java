package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Syncs {@code code_java/pom.xml} properties {@code jvm.max.heap} and {@code jvm.initial.heap} to the desired heap
 * size from the Memory Settings tab (MiB).
 */
public final class PomJvmHeapPropertiesSync {

    /** Minimum heap token for pom (2 GiB). */
    public static final int MIN_HEAP_MIB = 2048;

    private PomJvmHeapPropertiesSync() {}

    /**
     * Updates JVM heap properties in {@code code_java/pom.xml}.
     *
     * @param heapMaxMiB desired max heap MiB; values below {@link #MIN_HEAP_MIB} are clamped upward
     */
    public static void writeJvmHeapFromDesiredMiB(Map<String, String> ui, int heapMaxMiB) {
        int mib = Math.max(MIN_HEAP_MIB, heapMaxMiB);
        String token = formatJvmHeapToken(mib);
        Path root = AppPaths.resolveRepoRoot(ui != null ? ui : Map.of());
        Path pom = root.resolve("code_java").resolve("pom.xml");
        if (!Files.isRegularFile(pom)) {
            return;
        }
        String content;
        try {
            content = Files.readString(pom, StandardCharsets.UTF_8);
        } catch (IOException e) {
            System.err.println("[PM-AI] Failed to read code_java/pom.xml: " + pom + " : " + e.getMessage());
            return;
        }
        String updated = replaceProperty(content, "jvm.max.heap", token);
        updated = replaceProperty(updated, "jvm.initial.heap", token);
        if (updated.equals(content)) {
            return;
        }
        try {
            Files.writeString(pom, updated, StandardCharsets.UTF_8);
        } catch (IOException e) {
            System.err.println("[PM-AI] Failed to write heap props to code_java/pom.xml: " + pom + " : " + e.getMessage());
        }
    }

    /** Same convention as JVM flags: whole GiB as {@code Ng}, otherwise {@code Nm}. */
    static String formatJvmHeapToken(int mib) {
        if (mib >= 1024 && mib % 1024 == 0) {
            return (mib / 1024) + "g";
        }
        return mib + "m";
    }

    private static String replaceProperty(String xml, String propName, String value) {
        String escaped = propName.replace(".", "\\.");
        Pattern p = Pattern.compile("<" + escaped + ">[^<]*</" + escaped + ">");
        Matcher m = p.matcher(xml);
        String replacement =
                "<" + propName + ">" + Matcher.quoteReplacement(value) + "</" + propName + ">";
        return m.replaceFirst(replacement);
    }
}
