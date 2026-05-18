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
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.function.Consumer;

/**
 * Syncs portable bundle ({@code pm-ai-data}) from a canonical repo root at startup. Skips paths that must stay local
 * (output, selected workbooks, etc.).
 */
public final class PortableBundleSelfUpdater {

    /** Version-upgrade bundle file name on the release share ({@code pm-ai-package-release}). */
    public static final String PORTABLE_UPGRADE_ZIP_NAME = "PMD_version_upgrade.zip";

    /** Run-tab log prefix for portable bundle sync (filter-friendly). */
    public static final String PORTABLE_SYNC_LOG_PREFIX = "[portable-sync] ";

    private static final CopyOption[] COPY_OPTIONS =
            new CopyOption[] {StandardCopyOption.REPLACE_EXISTING, StandardCopyOption.COPY_ATTRIBUTES};

    private PortableBundleSelfUpdater() {}

    /** True when {@code pm-ai-data/code/python/task_extract_stage1.py} exists under {@code cwd}. */
    public static boolean isPortableBundleLayout(Path cwd) {
        Path marker =
                cwd.resolve("pm-ai-data").resolve("code").resolve("python").resolve("task_extract_stage1.py");
        return Files.isRegularFile(marker);
    }

    /**
     * Upgrade ZIP for sync: {@code canonical} when it is a {@code .zip}, or {@code canonical}/{@link
     * #PORTABLE_UPGRADE_ZIP_NAME} when that file exists under a release folder.
     */
    public static Optional<Path> resolveEffectiveUpgradeZip(Path canonical) {
        if (canonical == null) {
            return Optional.empty();
        }
        Path abs = canonical.toAbsolutePath().normalize();
        if (isPortableBundleZipPath(abs)) {
            return Optional.of(abs);
        }
        if (!isReadableDirectory(abs)) {
            return Optional.empty();
        }
        Path nested = abs.resolve(PORTABLE_UPGRADE_ZIP_NAME);
        if (isPortableBundleZipPath(nested)) {
            return Optional.of(nested);
        }
        return Optional.empty();
    }

    /**
     * Outer {@link AppPaths#VERSION_TXT_FILE_NAME} used for version compare and post-sync copy (beside upgrade ZIP or
     * directory root).
     */
    public static Optional<Path> resolveOuterVersionTxt(Path canonical) {
        if (canonical == null) {
            return Optional.empty();
        }
        Path abs = canonical.toAbsolutePath().normalize();
        Optional<Path> zip = resolveEffectiveUpgradeZip(abs);
        if (zip.isPresent()) {
            Path parent = zip.get().getParent();
            if (parent == null) {
                return Optional.empty();
            }
            return Optional.of(parent.resolve(AppPaths.VERSION_TXT_FILE_NAME));
        }
        Path atRoot = abs.resolve(AppPaths.VERSION_TXT_FILE_NAME);
        if (Files.isRegularFile(atRoot)) {
            return Optional.of(atRoot);
        }
        return Optional.empty();
    }

    /** Reads canonical version beside upgrade ZIP or at directory root. */
    public static Optional<BigDecimal> readCanonicalPortableBundleVersion(Path canonicalPath) {
        Objects.requireNonNull(canonicalPath, "canonicalPath");
        return resolveOuterVersionTxt(canonicalPath).flatMap(PortableBundleSelfUpdater::parseVersionFile);
    }

    /**
     * Directory whose tree is copied into local {@code pm-ai-data}: extracted ZIP {@code pm-ai-data}, nested {@code
     * pm-ai-data}, or repo-root layout (the canonical directory itself).
     */
    public static Path resolveSyncSourceRoot(Path canonical) {
        Objects.requireNonNull(canonical, "canonical");
        Path abs = canonical.toAbsolutePath().normalize();
        Path nested = abs.resolve("pm-ai-data");
        if (Files.isDirectory(nested)) {
            return nested;
        }
        return abs;
    }

    /** Reads {@link AppPaths#VERSION_TXT_FILE_NAME} under {@code pm-ai-data}, then {@code cwd} when missing (release ZIP omits inner version). */
    public static Optional<BigDecimal> readLocalBundleVersion(Path cwd, Path pmAiDataRoot) {
        Objects.requireNonNull(pmAiDataRoot, "pmAiDataRoot");
        Optional<BigDecimal> inData = parseVersionFile(pmAiDataRoot.resolve(AppPaths.VERSION_TXT_FILE_NAME));
        if (inData.isPresent()) {
            return inData;
        }
        if (cwd != null) {
            return parseVersionFile(cwd.toAbsolutePath().normalize().resolve(AppPaths.VERSION_TXT_FILE_NAME));
        }
        return Optional.empty();
    }

    /** {@code true} when {@code path} is a regular file whose name ends with {@code .zip} (case-insensitive). */
    public static boolean isPortableBundleZipPath(Path path) {
        if (path == null || !Files.isRegularFile(path)) {
            return false;
        }
        String name = path.getFileName().toString();
        return name.toLowerCase(Locale.ROOT).endsWith(".zip");
    }

    /** Readable directory, or readable {@code .zip} file for portable upgrade bundles. */
    public static boolean isValidPortableBundleCanonical(Path path) {
        if (path == null) {
            return false;
        }
        if (isPortableBundleZipPath(path)) {
            try {
                return Files.isReadable(path);
            } catch (SecurityException e) {
                return false;
            }
        }
        return isReadableDirectory(path);
    }

    /**
     * Extracts a portable app-image zip (root contains {@code pm-ai-data/}) to a new temp directory. Caller must
     * {@link #deleteDirectoryRecursive(Path, Consumer)} when done.
     */
    public static Path extractUpgradeZipToTempDirectory(Path zipPath, Consumer<String> log) throws IOException {
        Objects.requireNonNull(zipPath, "zipPath");
        Path zip = zipPath.toAbsolutePath().normalize();
        if (!isPortableBundleZipPath(zip)) {
            throw new IOException("Not a zip file: " + zip);
        }
        Path tempRoot = Files.createTempDirectory("pm-ai-upgrade-zip-");
        if (log != null) {
            log.accept(
                    PORTABLE_SYNC_LOG_PREFIX
                            + "ZIP 展開開始: "
                            + safePathForLog(zip)
                            + " -> "
                            + safePathForLog(tempRoot));
        }
        int[] extractedFiles = {0};
        try (java.util.zip.ZipFile zf = new java.util.zip.ZipFile(zip.toFile(), StandardCharsets.UTF_8)) {
            java.util.Enumeration<? extends java.util.zip.ZipEntry> en = zf.entries();
            while (en.hasMoreElements()) {
                java.util.zip.ZipEntry entry = en.nextElement();
                String name = entry.getName().replace('\\', '/');
                if (name.isEmpty() || name.startsWith("/") || name.contains("..")) {
                    throw new IOException("Unsafe or unsupported zip entry: " + name);
                }
                Path dest = tempRoot.resolve(name).normalize();
                if (!dest.startsWith(tempRoot)) {
                    throw new IOException("Zip-slip entry: " + name);
                }
                if (entry.isDirectory()) {
                    Files.createDirectories(dest);
                } else {
                    Files.createDirectories(dest.getParent());
                    try (var in = zf.getInputStream(entry)) {
                        Files.copy(in, dest, StandardCopyOption.REPLACE_EXISTING);
                    }
                    extractedFiles[0]++;
                    if (log != null) {
                        log.accept(PORTABLE_SYNC_LOG_PREFIX + "展開: " + name);
                    }
                }
            }
        }
        if (log != null) {
            log.accept(PORTABLE_SYNC_LOG_PREFIX + "ZIP 展開完了: ファイル " + extractedFiles[0] + " 件");
        }
        return tempRoot;
    }

    /** Best-effort recursive delete (reverse walk). */
    public static void deleteDirectoryRecursive(Path root, Consumer<String> log) {
        if (root == null) {
            return;
        }
        try {
            if (!Files.exists(root)) {
                return;
            }
            try (java.util.stream.Stream<Path> walk = Files.walk(root)) {
                walk.sorted(java.util.Comparator.reverseOrder())
                        .forEach(
                                p -> {
                                    try {
                                        Files.deleteIfExists(p);
                                    } catch (IOException e) {
                                        if (log != null) {
                                            log.accept("[portable-sync] cleanup: " + p + " — " + e.getMessage());
                                        }
                                    }
                                });
            }
        } catch (IOException e) {
            if (log != null) {
                log.accept("[portable-sync] cleanup walk failed: " + e.getMessage());
            }
        }
    }

    /** Reads {@link AppPaths#VERSION_TXT_FILE_NAME} at canonical repo root. */
    public static Optional<BigDecimal> readBundleVersion(Path canonicalRepoRoot) {
        Path v = canonicalRepoRoot.resolve(AppPaths.VERSION_TXT_FILE_NAME);
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
                        if (log != null) {
                            log.accept(
                                    PORTABLE_SYNC_LOG_PREFIX
                                            + "同期: "
                                            + rel.toString().replace('\\', '/'));
                        }
                        return FileVisitResult.CONTINUE;
                    }
                });

        if (log != null) {
            log.accept(
                    PORTABLE_SYNC_LOG_PREFIX
                            + "同期完了: "
                            + copied[0]
                            + " ファイル（除外パスはスキップ）");
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
        if (unix.equals("init_setting/user-profiles") || unix.startsWith("init_setting/user-profiles/")) {
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
        s.add("code/\u30b5\u30de\u30ea_AI\u914d\u53f0.xlsx");
        s.add("code/\u56fd\u5206\u30b5\u30de\u30ea_AI\u914d\u53f0.xlsx");
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

    /** Copies outer {@link AppPaths#VERSION_TXT_FILE_NAME} into {@code pmAiDataRoot} and {@code cwd} when present. */
    public static void copyOuterVersionTxtToLocal(Path canonical, Path cwd, Path pmAiDataRoot) throws IOException {
        Objects.requireNonNull(pmAiDataRoot, "pmAiDataRoot");
        Optional<Path> outer = resolveOuterVersionTxt(canonical);
        if (outer.isEmpty() || !Files.isRegularFile(outer.get())) {
            return;
        }
        Path src = outer.get();
        Files.createDirectories(pmAiDataRoot);
        Files.copy(src, pmAiDataRoot.resolve(AppPaths.VERSION_TXT_FILE_NAME), COPY_OPTIONS);
        if (cwd != null) {
            Files.copy(src, cwd.toAbsolutePath().normalize().resolve(AppPaths.VERSION_TXT_FILE_NAME), COPY_OPTIONS);
        }
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
