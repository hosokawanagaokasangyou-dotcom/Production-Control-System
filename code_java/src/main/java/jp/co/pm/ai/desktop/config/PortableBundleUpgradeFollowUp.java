package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.util.Objects;
import java.util.Optional;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * デスクトップ本体の終了後更新（{@link PortableBundleUpdateLauncher}）のあと、次回起動で工場既定選択などを続行するためのマーカー。
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public record PortableBundleUpgradeFollowUp(String installRoot, String version, String createdAt) {

    public static final String MANIFEST_FILE_NAME = "pending-portable-upgrade-followup.json";

    private static final ObjectMapper MAPPER = new ObjectMapper();

    public static Path manifestPath() {
        return PortableBundlePendingUpdate.userStateDirectory().resolve(MANIFEST_FILE_NAME);
    }

    public static void writePending(Path installRoot, String version) throws IOException {
        Objects.requireNonNull(installRoot, "installRoot");
        Files.createDirectories(PortableBundlePendingUpdate.userStateDirectory());
        PortableBundleUpgradeFollowUp state =
                new PortableBundleUpgradeFollowUp(
                        installRoot.toAbsolutePath().normalize().toString(),
                        version != null ? version.strip() : "",
                        Instant.now().toString());
        MAPPER.writerWithDefaultPrettyPrinter().writeValue(manifestPath().toFile(), state);
    }

    public static Optional<PortableBundleUpgradeFollowUp> readIfPresent() {
        Path file = manifestPath();
        if (!Files.isRegularFile(file)) {
            return Optional.empty();
        }
        try {
            return Optional.of(MAPPER.readValue(Files.readString(file), PortableBundleUpgradeFollowUp.class));
        } catch (Exception e) {
            return Optional.empty();
        }
    }

    public static boolean isPendingFor(Path installRoot) {
        if (installRoot == null) {
            return false;
        }
        return readIfPresent()
                .map(
                        f ->
                                Path.of(f.installRoot())
                                        .toAbsolutePath()
                                        .normalize()
                                        .equals(installRoot.toAbsolutePath().normalize()))
                .orElse(false);
    }

    public Path installRootPath() {
        return Path.of(installRoot).toAbsolutePath().normalize();
    }

    public static void clear() {
        try {
            Files.deleteIfExists(manifestPath());
        } catch (IOException ignored) {
            /* best-effort */
        }
    }
}
