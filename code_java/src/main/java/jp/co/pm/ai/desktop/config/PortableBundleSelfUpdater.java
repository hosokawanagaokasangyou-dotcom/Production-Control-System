package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.nio.file.CopyOption;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.StandardCopyOption;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.LinkedHashSet;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.function.Consumer;

/**
 * Syncs portable bundle ({@code pm-ai-data}) from a canonical repo root at startup. Skips paths that must stay local
 * (output, selected workbooks, etc.).
 */
public final class PortableBundleSelfUpdater {

    private static final CopyOption[] COPY_OPTIONS =
            new CopyOption[] {StandardCopyOption.REPLACE_EXISTING, StandardCopyOption.COPY_ATTRIBUTES};

    private PortableBundleSelfUpdater() {}

    /** True when {@code pm-ai-data/code/python/task_extract_stage1.py} exists under {@code cwd}. */
    public static boolean isPortableBundleLayout(Path cwd) {
        Path marker =
                cwd.resolve("pm-ai-data").resolve("code").resolve("python").resolve("task_extract_stage1.py");
        return Files.isRegularFile(marker);
    }

    /** Reads {@link AppPaths#VERSION_TXT_FILE_NAME} at canonical repo root. */
    public static Optional<BigDecimal> readBundleVersion(Path canonicalRepoRoot) {
        Path v = canonicalRepoRoot.resolve(AppPaths.VERSION_TXT_FILE_NAME);
        return parseVersionFile(v);
    }

    /** Reads {@link AppPaths#VERSION_TXT_FILE_NAME} under {@code pm-ai-data}. */
    public static Optional<BigDecimal> readLocalBundleVersion(Path pmAiDataRoot) {
        Path v = pmAiDataRoot.resolve(AppPaths.VERSION_TXT_FILE_NAME);
        return parseVersionFile(v);
    }

    /** Version used when local {@code version.txt} is missing (treated as oldest). */
    public static BigDecimal fallbackMinimumVersion() {
        return BigDecimal.ZERO;
    }

    /**
     * Copies files from {@code canonicalRepoRoot} into {@code localPmAiDataRoot}, skipping excluded relative paths.
     *
     * @return number of files copied (directories not counted).
     */
    public static SyncOutcome syncFromCanonical(Path canonicalRepoRoot, Path localPmAiDataRoot, Consumer<String> log)
            throws IOException {
        Objects.requireNonNull(canonicalRepoRoot, "canonicalRepoRoot");
        Objects.requireNonNull(localPmAiDataRoot, "localPmAiDataRoot");
        Path canon = canonicalRepoRoot.toAbsolutePath().normalize();
        Path destRoot = localPmAiDataRoot.toAbsolutePath().normalize();
        if (!Files.isDirectory(canon)) {
            throw new IOException("Canonical folder does not exist or is not a directory: " + canon);
        }
        Files.createDirectories(destRoot);

        int[] copied = {0};
        Files.walkFileTree(
                canon,
                new SimpleFileVisitor<>() {
                    @Override
                    public FileVisitResult preVisitDirectory(Path dir, BasicFileAttributes attrs) throws IOException {
                        Path rel = canon.relativize(dir);
                        if (rel.getNameCount() == 0) {
                            return FileVisitResult.CONTINUE;
                        }
                        if (isExcludedPath(rel)) {
                            return FileVisitResult.SKIP_SUBTREE;
                        }
                        Path targetDir = destRoot.resolve(rel.toString());
                        if (!Files.exists(targetDir)) {
                            Files.createDirectories(targetDir);
                        }
                        return FileVisitResult.CONTINUE;
                    }

                    @Override
                    public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
                        Path rel = canon.relativize(file);
                        if (isExcludedPath(rel)) {
                            return FileVisitResult.CONTINUE;
                        }
                        Path target = destRoot.resolve(rel.toString());
                        Files.createDirectories(target.getParent());
                        Files.copy(file, target, COPY_OPTIONS);
                        copied[0]++;
                        return FileVisitResult.CONTINUE;
                    }
                });

        if (log != null) {
            log.accept("[portable-sync] copied " + copied[0] + " files (exclusions skipped)");
        }
        return new SyncOutcome(copied[0]);
    }

    /** Package-private for tests: relative path from canonical repo root is excluded from sync. */
    static boolean isExcludedPath(Path relativeFromRepoRoot) {
        Path norm = relativeFromRepoRoot.normalize();
        String unix = norm.toString().replace('\\', '/');

        if (unix.isEmpty()) {
            return false;
        }
        if (unix.equals(".git") || unix.startsWith(".git/")) {
            return true;
        }
        if (unix.equals("output") || unix.startsWith("output/")) {
            return true;
        }
        if (unix.equals("input") || unix.startsWith("input/")) {
            return true;
        }
        if (unix.equals("code/log") || unix.startsWith("code/log/")) {
            return true;
        }
        if (unix.startsWith("code_java/build_cache/")
                || unix.startsWith("code_java/package_input/")
                || unix.startsWith("code_java/dist/")) {
            return true;
        }
        Set<String> exact = excludedExactRelativeUnixPaths();
        return exact.contains(unix);
    }

    /**
     * Exact repo-relative paths (forward slashes) that must never be overwritten from canonical. Japanese segments use
     * Unicode escapes so the source compiles regardless of host charset on CIFS mounts.
     */
    static Set<String> excludedExactRelativeUnixPaths() {
        LinkedHashSet<String> s = new LinkedHashSet<>();
        s.add("master.xlsm");
        s.add("code/\u30b5\u30de\u30ea_AI\u914d\u53f0.xlsm");
        s.add("code/\u7d50\u679c_\u914d\u53f0\u8868.json");
        s.add("code/\u7d50\u679c_\u914d\u53f0\u8868.xlsx");
        s.add("code/\u56fd\u5206master.xlsm");
        return s;
    }

    static Optional<BigDecimal> parseVersionFile(Path file) {
        if (!Files.isRegularFile(file)) {
            return Optional.empty();
        }
        try {
            String raw = Files.readString(file, StandardCharsets.UTF_8).trim();
            if (raw.isEmpty()) {
                return Optional.empty();
            }
            String firstLine = raw.lines().findFirst().orElse("").trim();
            if (firstLine.isEmpty()) {
                return Optional.empty();
            }
            return Optional.of(new BigDecimal(firstLine));
        } catch (Exception e) {
            return Optional.empty();
        }
    }

    /** Result of a sync run. */
    public record SyncOutcome(int filesCopied) {}

    /**
     * {@code true} when canonical {@link AppPaths#VERSION_TXT_FILE_NAME} is greater than local. If canonical has no
     * version file, do not sync.
     */
    public static boolean shouldUpdate(Optional<BigDecimal> canonicalVer, Optional<BigDecimal> localVer) {
        if (canonicalVer.isEmpty()) {
            return false;
        }
        BigDecimal remote = canonicalVer.get();
        BigDecimal local = localVer.orElse(fallbackMinimumVersion());
        return remote.compareTo(local) > 0;
    }

    /** Whether {@code p} is a readable directory (UNC-safe check). */
    public static boolean isReadableDirectory(Path p) {
        try {
            return Files.isDirectory(p) && Files.isReadable(p);
        } catch (SecurityException e) {
            return false;
        }
    }

    /** Safe absolute path string for logs. */
    public static String safePathForLog(Path p) {
        try {
            return Objects.requireNonNullElse(p, Path.of(".")).toAbsolutePath().normalize().toString();
        } catch (Exception e) {
            return "(unknown path)";
        }
    }
}
