package jp.co.pm.ai.kouchin.verify;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.concurrent.ConcurrentHashMap;

/**
 * UNC ディレクトリ一覧の mtime キャッシュ。再検出・再実行・CSV 取り込み後に {@link #invalidateAll()}。
 */
public final class KouchinDirListingCache {

    private record Entry(long mtime, List<Path> files) {}

    private static final ConcurrentHashMap<String, Entry> CACHE = new ConcurrentHashMap<>();

    private KouchinDirListingCache() {}

    public static void invalidateAll() {
        CACHE.clear();
    }

    public static List<Path> list(Path folder, String glob) {
        if (folder == null || !Files.isDirectory(folder)) {
            throw new VerifyException("フォルダにアクセスできません: " + folder);
        }
        long mtime = 0L;
        try {
            mtime = Files.getLastModifiedTime(folder).toMillis();
        } catch (Exception ignored) {
        }
        String key = folder.toAbsolutePath().normalize() + "\0" + glob;
        Entry cached = CACHE.get(key);
        if (cached != null && cached.mtime() == mtime) {
            return cached.files();
        }
        List<Path> files = FileDiscovery.listUncached(folder, glob);
        CACHE.put(key, new Entry(mtime, files));
        return files;
    }
}
