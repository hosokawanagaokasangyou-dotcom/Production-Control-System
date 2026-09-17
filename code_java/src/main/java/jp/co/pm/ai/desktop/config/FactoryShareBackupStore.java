package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.StandardCopyOption;
import java.nio.file.attribute.BasicFileAttributes;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Function;
import java.util.regex.Pattern;

/**
 * 湖南・国分の工場共有 DATA を、PC ローカルと各工場共有の {@code バックアップ/} へ世代コピーする。
 */
public final class FactoryShareBackupStore {

    public static final int MAX_GENERATIONS = 5;

    private static final DateTimeFormatter GEN_TS =
            DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss");
    private static final Pattern GEN_DIR = Pattern.compile("^\\d{8}-\\d{6}$");

    private FactoryShareBackupStore() {}

    public record SiteResult(
            FactorySite site,
            Path source,
            Path localDest,
            Path sharedDest,
            int copiedToLocal,
            int copiedToShared,
            String localError,
            String sharedError) {
        public boolean ok() {
            return (localError == null || localError.isBlank())
                    && (sharedError == null || sharedError.isBlank());
        }
    }

    public record Result(String generationId, Path localGenerationRoot, List<SiteResult> sites) {
        public boolean ok() {
            if (sites == null || sites.isEmpty()) {
                return false;
            }
            for (SiteResult s : sites) {
                if (!s.ok()) {
                    return false;
                }
            }
            return true;
        }
    }

    public static Path resolveLocalBackupRoot() {
        return AppPaths.resolveFactoryShareBackupLocalRoot();
    }

    public static Path resolveSharedBackupRoot(Map<String, String> ui, FactorySite site) {
        return AppPaths.resolveFactoryShareBackupSharedRoot(ui, site);
    }

    public static Path resolveSource(Map<String, String> ui, FactorySite site) {
        return AppPaths.summarySharedDataDirForFactory(ui, site);
    }

    public static boolean shouldSkipRelative(Path relative) {
        if (relative == null) {
            return false;
        }
        for (Path part : relative) {
            String name = part.toString();
            if (AppPaths.FACTORY_SHARE_BACKUP_DIR_NAME.equals(name)) {
                return true;
            }
            if (name.startsWith("~$")) {
                return true;
            }
        }
        return false;
    }

    public static String generationId(LocalDateTime when) {
        LocalDateTime t = when != null ? when : LocalDateTime.now(ZoneId.systemDefault());
        return GEN_TS.format(t);
    }

    /**
     * 湖南・国分の共有 DATA をローカルと各工場 {@code バックアップ/} へコピーする。
     */
    public static Result backupNow(Map<String, String> ui, LocalDateTime when) throws IOException {
        Map<FactorySite, Path> sources = new LinkedHashMap<>();
        for (FactorySite site : FactorySite.dispatchProductionSites()) {
            sources.put(site, resolveSource(ui, site));
        }
        return backup(
                sources,
                resolveLocalBackupRoot(),
                site -> resolveSharedBackupRoot(ui, site),
                when);
    }

    public static Result backup(
            Map<FactorySite, Path> sources,
            Path localRoot,
            Function<FactorySite, Path> sharedBackupRoot,
            LocalDateTime when)
            throws IOException {
        String gen = generationId(when);
        Path localGen = localRoot.resolve(gen);
        Files.createDirectories(localGen);
        List<SiteResult> sites = new ArrayList<>();
        for (FactorySite site : FactorySite.dispatchProductionSites()) {
            Path source = sources != null ? sources.get(site) : null;
            Path localDest = localGen.resolve(site.name());
            Path sharedRoot =
                    sharedBackupRoot != null ? sharedBackupRoot.apply(site) : null;
            Path sharedDest = sharedRoot != null ? sharedRoot.resolve(gen) : null;
            sites.add(backupOne(site, source, localDest, sharedDest));
            if (sharedRoot != null) {
                pruneOldGenerations(sharedRoot, MAX_GENERATIONS);
            }
        }
        pruneOldGenerations(localRoot, MAX_GENERATIONS);
        return new Result(gen, localGen.toAbsolutePath().normalize(), List.copyOf(sites));
    }

    private static SiteResult backupOne(
            FactorySite site, Path source, Path localDest, Path sharedDest) {
        int localCount = 0;
        int sharedCount = 0;
        String localError = "";
        String sharedError = "";
        if (source == null || !Files.isDirectory(source)) {
            localError = "ソースがありません";
            sharedError = localError;
            return new SiteResult(
                    site, source, localDest, sharedDest, 0, 0, localError, sharedError);
        }
        try {
            localCount = copyTree(source, localDest, true);
        } catch (IOException ex) {
            localError = ex.getMessage() != null ? ex.getMessage() : ex.toString();
        }
        try {
            if (sharedDest != null) {
                if (localError.isBlank() && Files.isDirectory(localDest)) {
                    sharedCount = copyTree(localDest, sharedDest, false);
                } else {
                    sharedCount = copyTree(source, sharedDest, true);
                }
            }
        } catch (IOException ex) {
            sharedError = ex.getMessage() != null ? ex.getMessage() : ex.toString();
        }
        return new SiteResult(
                site,
                source.toAbsolutePath().normalize(),
                localDest.toAbsolutePath().normalize(),
                sharedDest != null ? sharedDest.toAbsolutePath().normalize() : null,
                localCount,
                sharedCount,
                localError,
                sharedError);
    }

    static int copyTree(Path source, Path dest, boolean skipBackupFolder) throws IOException {
        Path src = source.toAbsolutePath().normalize();
        Path dst = dest.toAbsolutePath().normalize();
        Files.createDirectories(dst);
        int[] count = {0};
        Files.walkFileTree(
                src,
                new SimpleFileVisitor<>() {
                    @Override
                    public FileVisitResult preVisitDirectory(Path dir, BasicFileAttributes attrs)
                            throws IOException {
                        Path rel = src.relativize(dir);
                        if (!rel.toString().isEmpty()
                                && skipBackupFolder
                                && shouldSkipRelative(rel)) {
                            return FileVisitResult.SKIP_SUBTREE;
                        }
                        Files.createDirectories(dst.resolve(rel.toString()));
                        return FileVisitResult.CONTINUE;
                    }

                    @Override
                    public FileVisitResult visitFile(Path file, BasicFileAttributes attrs)
                            throws IOException {
                        Path rel = src.relativize(file);
                        if (skipBackupFolder && shouldSkipRelative(rel)) {
                            return FileVisitResult.CONTINUE;
                        }
                        Path target = dst.resolve(rel.toString());
                        Files.createDirectories(target.getParent());
                        Files.copy(file, target, StandardCopyOption.REPLACE_EXISTING);
                        count[0]++;
                        return FileVisitResult.CONTINUE;
                    }

                    @Override
                    public FileVisitResult visitFileFailed(Path file, IOException exc) {
                        return FileVisitResult.CONTINUE;
                    }
                });
        return count[0];
    }

    public static List<Path> pruneOldGenerations(Path root, int keep) throws IOException {
        List<Path> removed = new ArrayList<>();
        if (root == null || !Files.isDirectory(root) || keep < 0) {
            return removed;
        }
        List<Path> gens = new ArrayList<>();
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(root)) {
            for (Path child : stream) {
                if (Files.isDirectory(child)
                        && GEN_DIR.matcher(child.getFileName().toString()).matches()) {
                    gens.add(child);
                }
            }
        }
        gens.sort(Comparator.comparing((Path p) -> p.getFileName().toString()).reversed());
        for (int i = keep; i < gens.size(); i++) {
            deleteRecursively(gens.get(i));
            removed.add(gens.get(i).toAbsolutePath().normalize());
        }
        return removed;
    }

    private static void deleteRecursively(Path root) throws IOException {
        if (root == null || !Files.exists(root)) {
            return;
        }
        Files.walkFileTree(
                root,
                new SimpleFileVisitor<>() {
                    @Override
                    public FileVisitResult visitFile(Path file, BasicFileAttributes attrs)
                            throws IOException {
                        Files.deleteIfExists(file);
                        return FileVisitResult.CONTINUE;
                    }

                    @Override
                    public FileVisitResult postVisitDirectory(Path dir, IOException exc)
                            throws IOException {
                        Files.deleteIfExists(dir);
                        return FileVisitResult.CONTINUE;
                    }
                });
    }
}
