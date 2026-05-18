package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.MessageDigest;
import java.util.HexFormat;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * 配布正本の {@value #FILE_NAME}。{@code version.txt} が同じでもデスクトップ JAR が差し替わっているときに更新を検出する。
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public record PortableBundleBuildManifest(
        String version,
        String desktopJarFileName,
        long desktopJarSize,
        String desktopJarSha256) {

    public static final String FILE_NAME = "build-manifest.json";

    private static final ObjectMapper MAPPER = new ObjectMapper();

    /** {@link PortableBundleSelfUpdater#resolveOuterVersionTxt(Path)} と同じ親ディレクトリ。 */
    public static Optional<Path> resolveBesideCanonical(Path canonical) {
        return PortableBundleSelfUpdater.resolveOuterVersionTxt(canonical)
                .map(v -> v.getParent().resolve(FILE_NAME))
                .filter(Files::isRegularFile);
    }

    public static Optional<PortableBundleBuildManifest> read(Path manifestFile) {
        if (manifestFile == null || !Files.isRegularFile(manifestFile)) {
            return Optional.empty();
        }
        try {
            return Optional.of(MAPPER.readValue(Files.readString(manifestFile), PortableBundleBuildManifest.class));
        } catch (Exception e) {
            return Optional.empty();
        }
    }

    public static Optional<PortableBundleBuildManifest> readBesideCanonical(Path canonical) {
        return resolveBesideCanonical(canonical).flatMap(PortableBundleBuildManifest::read);
    }

    /** ローカル {@code app/} のデスクトップ JAR が正本 manifest と一致しない。 */
    public boolean desktopJarDiffersFromLocal(Path installRoot) {
        Objects.requireNonNull(installRoot, "installRoot");
        if (desktopJarFileName == null || desktopJarFileName.isBlank()) {
            return false;
        }
        Path localJar = installRoot.resolve("app").resolve(desktopJarFileName);
        if (!Files.isRegularFile(localJar)) {
            return true;
        }
        try {
            long size = Files.size(localJar);
            if (desktopJarSize > 0 && size != desktopJarSize) {
                return true;
            }
            if (desktopJarSha256 != null && !desktopJarSha256.isBlank()) {
                String localHash = sha256Hex(localJar).orElse("");
                return !desktopJarSha256.equalsIgnoreCase(localHash);
            }
        } catch (IOException e) {
            return true;
        }
        return false;
    }

    static Optional<String> sha256Hex(Path file) {
        try (InputStream in = Files.newInputStream(file)) {
            MessageDigest md = MessageDigest.getInstance("SHA-256");
            byte[] buf = new byte[64 * 1024];
            int n;
            while ((n = in.read(buf)) >= 0) {
                md.update(buf, 0, n);
            }
            return Optional.of(HexFormat.of().formatHex(md.digest()));
        } catch (Exception e) {
            return Optional.empty();
        }
    }

    /** manifest の版（あれば）を数値として解釈。 */
    public Optional<BigDecimal> versionAsDecimal() {
        if (version == null || version.isBlank()) {
            return Optional.empty();
        }
        try {
            String first = version.lines().findFirst().orElse("").trim();
            if (first.isEmpty()) {
                return Optional.empty();
            }
            return Optional.of(new BigDecimal(first));
        } catch (Exception e) {
            return Optional.empty();
        }
    }

    /** ログ向け要約（ASCII）。 */
    public String summaryForLog() {
        return "manifest version="
                + Objects.requireNonNullElse(version, "?")
                + " jar="
                + Objects.requireNonNullElse(desktopJarFileName, "?")
                + " size="
                + desktopJarSize
                + " sha256="
                + (desktopJarSha256 == null || desktopJarSha256.isBlank()
                        ? "?"
                        : desktopJarSha256.substring(0, Math.min(12, desktopJarSha256.length())) + "…");
    }
}
