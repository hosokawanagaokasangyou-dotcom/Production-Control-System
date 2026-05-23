package jp.co.pm.ai.desktop;

import java.util.Locale;

/** stdout/stderr one-line classification for run-tab log highlighting. */
enum LogLineKind {
    NORMAL,
    WARN,
    ERROR;

    static LogLineKind classify(String line) {
        if (line == null || line.isEmpty()) {
            return NORMAL;
        }
        if (isPortableSyncRoutineLine(line)) {
            return NORMAL;
        }
        String s = line.toLowerCase(Locale.ROOT);
        if (matchesError(s)) {
            return ERROR;
        }
        if (matchesWarn(s)) {
            return WARN;
        }
        return NORMAL;
    }

    /**
     * バージョンアップのファイル列挙行。パス中の {@code exception} / {@code error} / {@code warning} で誤判定しない。
     */
    private static boolean isPortableSyncRoutineLine(String line) {
        if (!line.contains("[portable-sync]")) {
            return false;
        }
        return line.contains("展開: ")
                || line.contains("同期: ")
                || line.contains("本体同期: ")
                || line.contains("ZIP 取得")
                || line.contains("ZIP 展開");
    }

    private static boolean matchesError(String s) {
        if (s.contains("[error]")
                || s.contains("traceback")
                || s.contains("planningvalidationerror")
                || s.contains("失敗")
                || s.contains("エラー")) {
            return true;
        }
        if (s.contains("exception")) {
            return true;
        }
        if (s.contains("fatal")) {
            return true;
        }
        return s.contains("error:") || s.contains("error :");
    }

    private static boolean matchesWarn(String s) {
        return s.contains("[warn")
                || s.contains("warning")
                || s.contains("warn:")
                || s.contains("警告")
                || s.contains("deprecated")
                || s.contains("userwarning");
    }
}
