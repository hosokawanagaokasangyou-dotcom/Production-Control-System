package jp.co.pm.ai.desktop.io.conflict;

import java.nio.file.Path;
import java.util.Map;

/** スナップショットの欠落（null）と空ファイル（length 0）を区別する。 */
final class SnapshotPresence {

    private SnapshotPresence() {}

    static byte[] get(Map<Path, byte[]> map, Path path) {
        if (map == null || !map.containsKey(path)) {
            return null;
        }
        return map.get(path);
    }

    static boolean isAbsent(byte[] bytes) {
        return bytes == null;
    }

    static boolean isPresent(byte[] bytes) {
        return bytes != null;
    }
}
