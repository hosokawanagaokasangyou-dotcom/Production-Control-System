package jp.co.pm.ai.desktop.io.conflict;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.HexFormat;

/** ファイル内容の SHA-256 指紋。欠落は {@link #ABSENT}。 */
public final class FileContentFingerprint {

    /** ファイルが存在しないときの sentinel（64 桁 hex と衝突しない固定文字列）。 */
    public static final String ABSENT = "ABSENT";

    private FileContentFingerprint() {}

    public static String sha256Hex(Path path) throws IOException {
        if (path == null || !Files.isRegularFile(path)) {
            return ABSENT;
        }
        return sha256Hex(Files.readAllBytes(path));
    }

    public static String sha256Hex(byte[] bytes) {
        if (bytes == null) {
            throw new IllegalArgumentException("bytes");
        }
        try {
            MessageDigest md = MessageDigest.getInstance("SHA-256");
            return HexFormat.of().formatHex(md.digest(bytes));
        } catch (NoSuchAlgorithmException e) {
            throw new IllegalStateException("SHA-256 unavailable", e);
        }
    }
}
