package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Syncs JVM heap size from the Memory Settings tab (MiB) to {@code code_java/pom.xml} and, when a packaged desktop
 * install is detected, to {@code launch-pm-ai-desktop.bat} and jpackage {@code app/*.cfg}.
 */
public final class PomJvmHeapPropertiesSync {

    /** Minimum heap token for pom (2 GiB). */
    public static final int MIN_HEAP_MIB = 2048;

    private static final Pattern XMS_FLAG = Pattern.compile("-Xms\\S+");
    private static final Pattern XMX_FLAG = Pattern.compile("-Xmx\\S+");

    private PomJvmHeapPropertiesSync() {}

    /**
     * Updates JVM heap properties in {@code code_java/pom.xml} and packaged launchers when present.
     *
     * @param heapMaxMiB desired max heap MiB; values below {@link #MIN_HEAP_MIB} are clamped upward
     */
    public static void writeJvmHeapFromDesiredMiB(Map<String, String> ui, int heapMaxMiB) {
        int mib = Math.max(MIN_HEAP_MIB, heapMaxMiB);
        String token = formatJvmHeapToken(mib);
        Map<String, String> env = ui != null ? ui : Map.of();
        syncPomXml(env, token);
        resolveDesktopInstallRoot(env).ifPresent(root -> syncPortableLaunchers(root, token));
    }

    /** Same convention as JVM flags: whole GiB as {@code Ng}, otherwise {@code Nm}. */
    static String formatJvmHeapToken(int mib) {
        if (mib >= 1024 && mib % 1024 == 0) {
            return (mib / 1024) + "g";
        }
        return mib + "m";
    }

    /** Replaces {@code -Xms}/{@code -Xmx} tokens in launcher text (BAT or jpackage cfg). */
    static String patchHeapFlagsInText(String content, String token) {
        if (content == null || content.isEmpty()) {
            return content;
        }
        String updated = XMS_FLAG.matcher(content).replaceAll("-Xms" + Matcher.quoteReplacement(token));
        return XMX_FLAG.matcher(updated).replaceAll("-Xmx" + Matcher.quoteReplacement(token));
    }

    static Optional<Path> resolveDesktopInstallRoot(Map<String, String> ui) {
        Path userDir = Path.of(System.getProperty("user.dir", ".")).toAbsolutePath().normalize();
        if (looksLikeDesktopInstallRoot(userDir)) {
            return Optional.of(userDir);
        }
        Path repo = AppPaths.resolveRepoRoot(ui != null ? ui : Map.of());
        if (repo.getFileName() != null && "pm-ai-data".equals(repo.getFileName().toString())) {
            Path parent = repo.getParent();
            if (parent != null && looksLikeDesktopInstallRoot(parent)) {
                return Optional.of(parent);
            }
        }
        return Optional.empty();
    }

    static boolean looksLikeDesktopInstallRoot(Path dir) {
        if (dir == null || !Files.isDirectory(dir)) {
            return false;
        }
        if (Files.isRegularFile(dir.resolve(PortableBundleSelfUpdater.PORTABLE_DESKTOP_EXE_NAME))) {
            return true;
        }
        if (Files.isRegularFile(dir.resolve(PortableBundleSelfUpdater.PORTABLE_LAUNCHER_BAT_NAME))) {
            return true;
        }
        return Files.isDirectory(dir.resolve("app")) && Files.isDirectory(dir.resolve("runtime"));
    }

    private static void syncPomXml(Map<String, String> ui, String token) {
        Path root = AppPaths.resolveRepoRoot(ui);
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

    private static void syncPortableLaunchers(Path installRoot, String token) {
        Path bat = installRoot.resolve(PortableBundleSelfUpdater.PORTABLE_LAUNCHER_BAT_NAME);
        patchLauncherFileIfPresent(bat, token);

        Path appDir = installRoot.resolve("app");
        if (!Files.isDirectory(appDir)) {
            return;
        }
        Path namedCfg = appDir.resolve("PMD.cfg");
        if (Files.isRegularFile(namedCfg)) {
            patchLauncherFileIfPresent(namedCfg, token);
            return;
        }
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(appDir, "*.cfg")) {
            for (Path cfg : stream) {
                patchLauncherFileIfPresent(cfg, token);
            }
        } catch (IOException e) {
            System.err.println(
                    "[PM-AI] Failed to scan jpackage cfg under "
                            + appDir
                            + " : "
                            + e.getMessage());
        }
    }

    private static void patchLauncherFileIfPresent(Path file, String token) {
        if (!Files.isRegularFile(file)) {
            return;
        }
        String content;
        try {
            content = Files.readString(file, StandardCharsets.UTF_8);
        } catch (IOException e) {
            System.err.println("[PM-AI] Failed to read launcher file " + file + " : " + e.getMessage());
            return;
        }
        String updated = patchHeapFlagsInText(content, token);
        if (updated.equals(content)) {
            return;
        }
        try {
            Files.writeString(file, updated, StandardCharsets.UTF_8);
        } catch (IOException e) {
            System.err.println("[PM-AI] Failed to write launcher file " + file + " : " + e.getMessage());
        }
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
