package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * メモリ設定タブの希望ヒープ（MiB）に合わせて {@code code_java/pom.xml} の {@code jvm.max.heap} と
 * {@code jvm.initial.heap} を同一値に更新する。
 */
public final class PomJvmHeapPropertiesSync {

    /** pom のヒープ指定の下限（2 GiB）。 */
    public static final int MIN_HEAP_MIB = 2048;

    private PomJvmHeapPropertiesSync() {}

    /**
     * {@code code_java/pom.xml} の JVM ヒーププロパティを更新する。
     *
     * @param heapMaxMiB 希望上限（MiB）。{@code MIN_HEAP_MIB} 未満は {@code MIN_HEAP_MIB} に切り上げる。
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
            System.err.println(
                    "[PM-AI] code_java/pom.xml を読み込めませんでした: " + pom + " ? " + e.getMessage());
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
            System.err.println(
                    "[PM-AI] code_java/pom.xml にヒープ設定を書き込めませんでした: " + pom + " ? " + e.getMessage());
        }
    }

    /** JVM オプションと同様の表記（整数 GiB は {@code Ng}、それ以外は {@code Nm}）。 */
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
