package jp.co.pm.ai.desktop.io.conflict;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/** 読込時のパスごとの SHA-256 と内容スナップショット。 */
public final class FingerprintBaseline {

    private final Map<Path, String> hashes;
    private final Map<Path, byte[]> snapshots;

    private FingerprintBaseline(Map<Path, String> hashes, Map<Path, byte[]> snapshots) {
        this.hashes = Collections.unmodifiableMap(new LinkedHashMap<>(hashes));
        this.snapshots = Collections.unmodifiableMap(new LinkedHashMap<>(snapshots));
    }

    public static FingerprintBaseline capture(List<Path> paths) throws IOException {
        Objects.requireNonNull(paths, "paths");
        Map<Path, String> hashes = new LinkedHashMap<>();
        Map<Path, byte[]> snapshots = new LinkedHashMap<>();
        for (Path p : paths) {
            Path abs = p.toAbsolutePath().normalize();
            if (Files.isRegularFile(abs)) {
                byte[] bytes = Files.readAllBytes(abs);
                snapshots.put(abs, bytes);
                hashes.put(abs, FileContentFingerprint.sha256Hex(bytes));
            } else {
                snapshots.put(abs, new byte[0]);
                hashes.put(abs, FileContentFingerprint.ABSENT);
            }
        }
        return new FingerprintBaseline(hashes, snapshots);
    }

    public Map<Path, String> hashes() {
        return hashes;
    }

    public Map<Path, byte[]> snapshots() {
        return snapshots;
    }

    public List<Path> paths() {
        return List.copyOf(hashes.keySet());
    }
}
