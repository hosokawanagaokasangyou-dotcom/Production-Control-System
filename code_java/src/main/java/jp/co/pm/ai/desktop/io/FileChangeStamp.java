package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.NoSuchFileException;
import java.nio.file.Path;
import java.nio.file.attribute.BasicFileAttributes;

/**
 * ポーリング監視用の軽量ファイル指紋（パス・更新時刻・サイズ）。内容は読まない。
 *
 * <p>欠落は {@code exists=false}（時刻・サイズ 0）。共有フォルダの一時的な読取失敗は欠落と区別するため
 * {@link #read} が {@code null} を返す。
 */
public record FileChangeStamp(Path path, boolean exists, long lastModifiedMillis, long size) {

    /** @return 指紋。属性を取得できないとき（欠落以外の I/O 失敗）は {@code null} */
    public static FileChangeStamp read(Path path) {
        Path abs = path.toAbsolutePath().normalize();
        try {
            BasicFileAttributes attrs = Files.readAttributes(abs, BasicFileAttributes.class);
            if (!attrs.isRegularFile()) {
                return absent(abs);
            }
            return new FileChangeStamp(
                    abs, true, attrs.lastModifiedTime().toMillis(), attrs.size());
        } catch (NoSuchFileException e) {
            return absent(abs);
        } catch (IOException | SecurityException e) {
            return null;
        }
    }

    private static FileChangeStamp absent(Path abs) {
        return new FileChangeStamp(abs, false, 0L, 0L);
    }
}
