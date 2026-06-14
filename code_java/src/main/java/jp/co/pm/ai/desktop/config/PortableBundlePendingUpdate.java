package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
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
        return userStateDirectory(PortableBundleProfile.PMD);
    }

    public static Path userStateDirectory(PortableBundleProfile profile) {
        return Path.of(System.getProperty("user.home", "."), profile.userStateDirName());
    }

    public static Path manifestPath() {
        return manifestPath(PortableBundleProfile.PMD);
    }

    public static Path manifestPath(PortableBundleProfile profile) {
        return userStateDirectory(profile).resolve(profile.pendingManifestFileName());
    }

    public static Path defaultStagingDirectory() {
        return defaultStagingDirectory(PortableBundleProfile.PMD);
    }

    public static Path defaultStagingDirectory(PortableBundleProfile profile) {
        return userStateDirectory(profile).resolve(profile.stagingDirName());
    }

    public static void write(
            PortableBundleProfile profile,
            Path installRoot,
            Path stagingRoot,
            String version,
            long waitPid,
            Path canonicalPath)
            throws IOException {
        Objects.requireNonNull(profile, "profile");
        Objects.requireNonNull(installRoot, "installRoot");
        Objects.requireNonNull(stagingRoot, "stagingRoot");
        Files.createDirectories(userStateDirectory(profile));
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
        MAPPER.writerWithDefaultPrettyPrinter().writeValue(manifestPath(profile).toFile(), pending);
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

    public static Optional<PortableBundlePendingUpdate> readIfPresent(PortableBundleProfile profile) {
        Path file = manifestPath(profile);
        if (!Files.isRegularFile(file)) {
            return Optional.empty();
        }
        try {
            return Optional.of(MAPPER.readValue(Files.readString(file), PortableBundlePendingUpdate.class));
        } catch (Exception e) {
            return Optional.empty();
        }
    }

    public static Optional<PortableBundlePendingUpdate> readIfPresent() {
        return readIfPresent(PortableBundleProfile.PMD);
    }

    public static void clear(PortableBundleProfile profile) {
        try {
            Files.deleteIfExists(manifestPath(profile));
        } catch (IOException ignored) {
            /* best-effort */
        }
    }

    public static void clear() {
        clear(PortableBundleProfile.PMD);
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
