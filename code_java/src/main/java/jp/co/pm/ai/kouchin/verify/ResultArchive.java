package jp.co.pm.ai.kouchin.verify;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.stream.Stream;

/**
 * 最新以外の結果を {@code 過去検証結果} / {@code 過去月トレンド} へ退避する。
 * keep はその出力ディレクトリ上のパスのみ許可する。
 */
public final class ResultArchive {

    public static final String PAST_VERIFY = "過去検証結果";
    public static final String PAST_TREND = "過去月トレンド";

    private ResultArchive() {}

    public static void archiveOldVerifyResults(Path outputDir, Collection<Path> keepInThisDir)
            throws IOException {
        archiveMatching(outputDir, PAST_VERIFY, "検証結果_", keepInThisDir);
        archiveMatching(outputDir, PAST_VERIFY, "報告メール_統合_", keepInThisDir);
    }

    public static void archiveOldTrends(Path outputDir, Collection<Path> keepInThisDir)
            throws IOException {
        archiveMatching(outputDir, PAST_TREND, "月トレンド_", keepInThisDir);
    }

    static void archiveMatching(
            Path outputDir, String pastFolderName, String namePrefix, Collection<Path> keepInThisDir)
            throws IOException {
        if (outputDir == null || !Files.isDirectory(outputDir)) {
            return;
        }
        Path base = outputDir.toAbsolutePath().normalize();
        Set<Path> keep = new HashSet<>();
        if (keepInThisDir != null) {
            for (Path k : keepInThisDir) {
                if (k == null) {
                    continue;
                }
                Path n = k.toAbsolutePath().normalize();
                if (n.startsWith(base)) {
                    keep.add(n);
                }
            }
        }
        Path past = base.resolve(pastFolderName);
        List<Path> candidates = new ArrayList<>();
        try (Stream<Path> stream = Files.list(base)) {
            stream.filter(Files::isRegularFile)
                    .filter(p -> {
                        String n = p.getFileName().toString();
                        return n.startsWith(namePrefix)
                                && (n.endsWith(".xlsx")
                                        || n.endsWith(".html")
                                        || n.endsWith(".txt"));
                    })
                    .forEach(p -> candidates.add(p.toAbsolutePath().normalize()));
        }
        for (Path file : candidates) {
            if (keep.contains(file)) {
                continue;
            }
            Files.createDirectories(past);
            Path dest = uniqueDest(past, file.getFileName().toString());
            Files.move(file, dest, StandardCopyOption.REPLACE_EXISTING);
        }
    }

    static Path uniqueDest(Path pastDir, String fileName) throws IOException {
        Path dest = pastDir.resolve(fileName);
        if (!Files.exists(dest)) {
            return dest;
        }
        String stem;
        String ext;
        int dot = fileName.lastIndexOf('.');
        if (dot > 0) {
            stem = fileName.substring(0, dot);
            ext = fileName.substring(dot);
        } else {
            stem = fileName;
            ext = "";
        }
        int i = 1;
        while (Files.exists(dest)) {
            dest = pastDir.resolve(stem + "_dup" + i + ext);
            i++;
        }
        return dest;
    }
}
