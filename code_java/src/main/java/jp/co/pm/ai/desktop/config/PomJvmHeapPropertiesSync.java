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
     * Updates JVM heap properties from memory settings (fixed or variable mode).
     *
     * @param heapFixed {@code true}: both -Xms/-Xmx use {@code heapMinMiB}/{@code heapMaxMiB}; {@code false}: -Xms is
     *     {@link #MIN_HEAP_MIB} and -Xmx is {@code heapMaxMiB}
     * @param heapMinMiB desired min heap MiB (fixed mode only)
     * @param heapMaxMiB desired max heap MiB; values below {@link #MIN_HEAP_MIB} are clamped upward
     */
    public static void writeJvmHeapFromLaunchPrefs(
            Map<String, String> ui, boolean heapFixed, int heapMinMiB, int heapMaxMiB) {
        int maxMib = Math.max(MIN_HEAP_MIB, heapMaxMiB);
        int minMib =
                heapFixed
                        ? Math.max(MIN_HEAP_MIB, Math.min(Math.max(heapMinMiB, MIN_HEAP_MIB), maxMib))
                        : MIN_HEAP_MIB;
        String minToken = formatJvmHeapToken(minMib);
        String maxToken = formatJvmHeapToken(maxMib);
        Map<String, String> env = ui != null ? ui : Map.of();
        syncPomXml(env, minToken, maxToken);
        resolveDesktopInstallRoot(env).ifPresent(root -> syncPortableLaunchers(root, minToken, maxToken));
    }

    /** @deprecated use {@link #writeJvmHeapFromLaunchPrefs} */
    @Deprecated
    public static void writeJvmHeapFromDesiredMiB(Map<String, String> ui, int heapMaxMiB) {
        writeJvmHeapFromLaunchPrefs(ui, true, heapMaxMiB, heapMaxMiB);
    }

    /** Same convention as JVM flags: whole GiB as {@code Ng}, otherwise {@code Nm}. */
    static String formatJvmHeapToken(int mib) {
        if (mib >= 1024 && mib % 1024 == 0) {
            return (mib / 1024) + "g";
        }
        return mib + "m";
    }

    /** Replaces {@code -Xms}/{@code -Xmx} tokens in launcher text (BAT or jpackage cfg). */
    static String patchHeapFlagsInText(String content, String minToken, String maxToken) {
        if (content == null || content.isEmpty()) {
            return content;
        }
        String updated = XMS_FLAG.matcher(content).replaceAll("-Xms" + Matcher.quoteReplacement(minToken));
        return XMX_FLAG.matcher(updated).replaceAll("-Xmx" + Matcher.quoteReplacement(maxToken));
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

    private static void syncPomXml(Map<String, String> ui, String minToken, String maxToken) {
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
        String updated = replaceProperty(content, "jvm.max.heap", maxToken);
        updated = replaceProperty(updated, "jvm.initial.heap", minToken);
        if (updated.equals(content)) {
            return;
        }
        try {
            Files.writeString(pom, updated, StandardCharsets.UTF_8);
        } catch (IOException e) {
            System.err.println("[PM-AI] Failed to write heap props to code_java/pom.xml: " + pom + " : " + e.getMessage());
        }
    }

    private static void syncPortableLaunchers(Path installRoot, String minToken, String maxToken) {
        Path bat = installRoot.resolve(PortableBundleSelfUpdater.PORTABLE_LAUNCHER_BAT_NAME);
        patchLauncherFileIfPresent(bat, minToken, maxToken);

        Path appDir = installRoot.resolve("app");
        if (!Files.isDirectory(appDir)) {
            return;
        }
        Path namedCfg = appDir.resolve("PMD.cfg");
        if (Files.isRegularFile(namedCfg)) {
            patchLauncherFileIfPresent(namedCfg, minToken, maxToken);
            return;
        }
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(appDir, "*.cfg")) {
            for (Path cfg : stream) {
                patchLauncherFileIfPresent(cfg, minToken, maxToken);
            }
        } catch (IOException e) {
            System.err.println(
                    "[PM-AI] Failed to scan jpackage cfg under "
                            + appDir
                            + " : "
                            + e.getMessage());
        }
    }

    private static void patchLauncherFileIfPresent(Path file, String minToken, String maxToken) {
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
        String updated = patchHeapFlagsInText(content, minToken, maxToken);
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
