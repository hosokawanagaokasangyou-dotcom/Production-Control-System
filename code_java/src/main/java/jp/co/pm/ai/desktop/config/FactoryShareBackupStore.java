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
            return localError == null && sharedError == null;
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

        /** 失敗した工場と行き先。成功時は空文字。 */
        public String failureSummaryJa() {
            if (sites == null || sites.isEmpty()) {
                return "結果が空です";
            }
            List<String> parts = new ArrayList<>();
            for (SiteResult s : sites) {
                if (s == null || s.ok()) {
                    continue;
                }
                String who = s.site() != null ? s.site().displayLabelJa() : "工場";
                if (s.localError() != null) {
                    parts.add(who + " ローカル");
                }
                if (s.sharedError() != null) {
                    parts.add(who + " 共有");
                }
            }
            return String.join("、", parts);
        }
    }

    public static Path resolveLocalBackupRoot() {
        return AppPaths.resolveFactoryShareBackupLocalRoot();
    }

    public static Path resolveSharedBackupRoot(Map<String, String> ui, FactorySite site) {
        return AppPaths.resolveFactoryShareBackupSharedRoot(ui, site);
    }

    /**
     * 対象工場の共有 DATA。現在 UI の上書きが工場ヒント無しで両サイト同一になるときは、
     * 他工場だけ工場既定 UNC へ落とす。
     */
    public static Path resolveSource(Map<String, String> ui, FactorySite site) {
        Path resolved = AppPaths.summarySharedDataDirForFactory(ui, site);
        if (site == null || resolved == null) {
            return resolved;
        }
        FactorySite current = AppPaths.currentDispatchFactorySite(ui);
        if (current == site) {
            return resolved;
        }
        Path currentPath = AppPaths.summarySharedDataDirForFactory(ui, current);
        if (currentPath == null || !resolved.equals(currentPath)) {
            return resolved;
        }
        String factoryDefault = site.pmAiSummaryAiDispatchWorkbookEnvValue(ui);
        if (factoryDefault == null || factoryDefault.isBlank()) {
            return resolved;
        }
        Path fallback = Path.of(factoryDefault.trim()).toAbsolutePath().normalize();
        if (Files.isRegularFile(fallback)
                && fallback.getFileName() != null
                && fallback.getFileName().toString().toLowerCase().endsWith(".xlsx")) {
            Path parent = fallback.getParent();
            return parent != null ? parent : fallback;
        }
        return fallback;
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
        Files.createDirectories(localRoot);
        String gen = allocateGenerationId(localRoot, sharedBackupRoot, when);
        Path localGen = localRoot.resolve(gen);
        Files.createDirectories(localGen);
        List<SiteResult> sites = new ArrayList<>();
        boolean anyLocalOk = false;
        for (FactorySite site : FactorySite.dispatchProductionSites()) {
            Path source = sources != null ? sources.get(site) : null;
            Path localDest = localGen.resolve(site.name());
            Path sharedRoot =
                    sharedBackupRoot != null ? sharedBackupRoot.apply(site) : null;
            Path sharedDest =
                    sharedRoot != null ? sharedRoot.resolve(gen).resolve(site.name()) : null;
            SiteResult one = backupOne(site, source, localDest, sharedDest);
            if (one.localError() != null) {
                deleteQuietly(localDest);
            } else {
                anyLocalOk = true;
            }
            if (one.sharedError() != null) {
                deleteQuietly(sharedDest);
            } else if (sharedRoot != null) {
                pruneQuietly(sharedRoot);
            }
            sites.add(one);
        }
        if (anyLocalOk) {
            pruneQuietly(localRoot);
        } else {
            deleteQuietly(localGen);
        }
        Path localGenAbs = localGen.toAbsolutePath().normalize();
        return new Result(gen, localGenAbs, List.copyOf(sites));
    }

    static String allocateGenerationId(
            Path localRoot, Function<FactorySite, Path> sharedBackupRoot, LocalDateTime when) {
        LocalDateTime base = when != null ? when : LocalDateTime.now(ZoneId.systemDefault());
        for (int i = 0; i < 120; i++) {
            String gen = generationId(base.plusSeconds(i));
            if (generationExists(localRoot, sharedBackupRoot, gen)) {
                continue;
            }
            return gen;
        }
        return generationId(base) + "-" + Long.toHexString(System.nanoTime());
    }

    private static boolean generationExists(
            Path localRoot, Function<FactorySite, Path> sharedBackupRoot, String gen) {
        if (localRoot != null && Files.exists(localRoot.resolve(gen))) {
            return true;
        }
        if (sharedBackupRoot == null) {
            return false;
        }
        for (FactorySite site : FactorySite.dispatchProductionSites()) {
            Path sharedRoot = sharedBackupRoot.apply(site);
            if (sharedRoot != null && Files.exists(sharedRoot.resolve(gen))) {
                return true;
            }
        }
        return false;
    }

    private static SiteResult backupOne(
            FactorySite site, Path source, Path localDest, Path sharedDest) {
        int localCount = 0;
        int sharedCount = 0;
        String localError = null;
        String sharedError = null;
        if (source == null || !Files.isDirectory(source)) {
            localError = "ソースがありません";
            sharedError = localError;
            return new SiteResult(
                    site, source, localDest, sharedDest, 0, 0, localError, sharedError);
        }
        try {
            localCount = copyTree(source, localDest, true);
        } catch (IOException ex) {
            localError = errorText(ex);
        }
        try {
            if (sharedDest == null) {
                sharedError = "共有先がありません";
            } else if (localError == null && Files.isDirectory(localDest)) {
                sharedCount = copyTree(localDest, sharedDest, false);
            } else {
                sharedCount = copyTree(source, sharedDest, true);
            }
        } catch (IOException ex) {
            sharedError = errorText(ex);
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

    private static String errorText(Exception ex) {
        if (ex == null) {
            return "失敗";
        }
        String msg = ex.getMessage();
        if (msg != null && !msg.isBlank()) {
            return msg;
        }
        return ex.toString();
    }

    static int copyTree(Path source, Path dest, boolean skipBackupFolder) throws IOException {
        Path src = source.toAbsolutePath().normalize();
        Path dst = dest.toAbsolutePath().normalize();
        Files.createDirectories(dst);
        int[] count = {0};
        List<String> failures = new ArrayList<>();
        Files.walkFileTree(
                src,
                new SimpleFileVisitor<>() {
                    @Override
                    public FileVisitResult preVisitDirectory(Path dir, BasicFileAttributes attrs)
                            throws IOException {
                        Path abs = dir.toAbsolutePath().normalize();
                        if (!abs.equals(src) && isUnderOrSame(abs, dst)) {
                            return FileVisitResult.SKIP_SUBTREE;
                        }
                        Path rel = src.relativize(dir);
                        if (!rel.toString().isEmpty()
                                && skipBackupFolder
                                && shouldSkipRelative(rel)) {
                            return FileVisitResult.SKIP_SUBTREE;
                        }
                        Files.createDirectories(resolveUnder(dst, rel));
                        return FileVisitResult.CONTINUE;
                    }

                    @Override
                    public FileVisitResult visitFile(Path file, BasicFileAttributes attrs)
                            throws IOException {
                        Path abs = file.toAbsolutePath().normalize();
                        if (isUnderOrSame(abs, dst)) {
                            return FileVisitResult.CONTINUE;
                        }
                        Path rel = src.relativize(file);
                        if (skipBackupFolder && shouldSkipRelative(rel)) {
                            return FileVisitResult.CONTINUE;
                        }
                        Path target = resolveUnder(dst, rel);
                        Path parent = target.getParent();
                        if (parent != null) {
                            Files.createDirectories(parent);
                        }
                        Files.copy(file, target, StandardCopyOption.REPLACE_EXISTING);
                        count[0]++;
                        return FileVisitResult.CONTINUE;
                    }

                    @Override
                    public FileVisitResult visitFileFailed(Path file, IOException exc) {
                        Path abs = file != null ? file.toAbsolutePath().normalize() : null;
                        if (abs != null && isUnderOrSame(abs, dst)) {
                            return FileVisitResult.CONTINUE;
                        }
                        failures.add(
                                (file != null ? file : Path.of("?"))
                                        + ": "
                                        + errorText(exc));
                        return FileVisitResult.CONTINUE;
                    }
                });
        if (!failures.isEmpty()) {
            throw new IOException("読めなかったファイル: " + String.join("; ", failures));
        }
        return count[0];
    }

    private static Path resolveUnder(Path dst, Path rel) throws IOException {
        Path target = (rel == null || rel.toString().isEmpty()) ? dst : dst.resolve(rel);
        Path norm = target.toAbsolutePath().normalize();
        if (!isUnderOrSame(norm, dst)) {
            throw new IOException("コピー先がバックアップディレクトリの外です: " + norm);
        }
        return norm;
    }

    private static boolean isUnderOrSame(Path candidate, Path root) {
        if (candidate == null || root == null) {
            return false;
        }
        Path c = candidate.toAbsolutePath().normalize();
        Path r = root.toAbsolutePath().normalize();
        return c.equals(r) || c.startsWith(r);
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

    private static void pruneQuietly(Path root) {
        try {
            pruneOldGenerations(root, MAX_GENERATIONS);
        } catch (IOException ignored) {
            // コピー済み世代は残し、他工場の処理は続ける
        }
    }

    private static void deleteQuietly(Path root) {
        try {
            deleteRecursively(root);
        } catch (IOException ignored) {
            // 失敗世代の掃除に失敗してもバックアップ結果は既に記録済み
        }
    }

    private static void deleteRecursively(Path root) throws IOException {
        if (root == null || !Files.exists(root)) {
            return;
        }
        List<String> failures = new ArrayList<>();
        Files.walkFileTree(
                root,
                new SimpleFileVisitor<>() {
                    @Override
                    public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) {
                        try {
                            Files.deleteIfExists(file);
                        } catch (IOException ex) {
                            failures.add(file + ": " + errorText(ex));
                        }
                        return FileVisitResult.CONTINUE;
                    }

                    @Override
                    public FileVisitResult visitFileFailed(Path file, IOException exc) {
                        failures.add((file != null ? file : Path.of("?")) + ": " + errorText(exc));
                        return FileVisitResult.CONTINUE;
                    }

                    @Override
                    public FileVisitResult postVisitDirectory(Path dir, IOException exc) {
                        try {
                            Files.deleteIfExists(dir);
                        } catch (IOException ex) {
                            failures.add(dir + ": " + errorText(ex));
                        }
                        return FileVisitResult.CONTINUE;
                    }
                });
        if (!failures.isEmpty()) {
            throw new IOException("削除できませんでした: " + String.join("; ", failures));
        }
    }
}
