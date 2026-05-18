package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.Instant;
import java.util.Objects;
import java.util.Optional;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.databind.ObjectMapper;

/** 終了後にデスクトップ本体（{@code PMD.exe} / {@code app} / {@code runtime}）を適用するための状態。 */
@JsonIgnoreProperties(ignoreUnknown = true)
public record PortableBundlePendingUpdate(
        String version,
        String installRoot,
        String stagingRoot,
        long waitPid,
        String canonicalPath,
        String createdAt) {

    public static final String MANIFEST_FILE_NAME = "pending-portable-update.json";

    private static final ObjectMapper MAPPER = new ObjectMapper();

    public static Path userStateDirectory() {
        return Paths.get(System.getProperty("user.home", "."), ".pm-ai-desktop");
    }

    public static Path manifestPath() {
        return userStateDirectory().resolve(MANIFEST_FILE_NAME);
    }

    public static Path defaultStagingDirectory() {
        return userStateDirectory().resolve("pending-portable-update-staging");
    }

    public static void write(
            Path installRoot,
            Path stagingRoot,
            String version,
            long waitPid,
            Path canonicalPath)
            throws IOException {
        Objects.requireNonNull(installRoot, "installRoot");
        Objects.requireNonNull(stagingRoot, "stagingRoot");
        Files.createDirectories(userStateDirectory());
        PortableBundlePendingUpdate pending =
                new PortableBundlePendingUpdate(
                        version,
                        installRoot.toAbsolutePath().normalize().toString(),
                        stagingRoot.toAbsolutePath().normalize().toString(),
                        waitPid,
                        canonicalPath != null
                                ? canonicalPath.toAbsolutePath().normalize().toString()
                                : null,
                        Instant.now().toString());
        MAPPER.writerWithDefaultPrettyPrinter().writeValue(manifestPath().toFile(), pending);
    }

    public static Optional<PortableBundlePendingUpdate> readIfPresent() {
        Path file = manifestPath();
        if (!Files.isRegularFile(file)) {
            return Optional.empty();
        }
        try {
            return Optional.of(MAPPER.readValue(Files.readString(file), PortableBundlePendingUpdate.class));
        } catch (Exception e) {
            return Optional.empty();
        }
    }

    public static void clear() {
        try {
            Files.deleteIfExists(manifestPath());
        } catch (IOException ignored) {
            /* best-effort */
        }
    }

    public static void clearStaging(Path stagingRoot) {
        PortableBundleSelfUpdater.deleteDirectoryRecursive(stagingRoot, null);
    }

    public Path installRootPath() {
        return Path.of(installRoot).toAbsolutePath().normalize();
    }

    public Path stagingRootPath() {
        return Path.of(stagingRoot).toAbsolutePath().normalize();
    }
}
