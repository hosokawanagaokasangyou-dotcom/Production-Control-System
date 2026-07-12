package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.nio.file.CopyOption;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.security.MessageDigest;
import java.util.HexFormat;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.function.Consumer;

import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * 同梱 {@link AppPaths#RDP_LAUNCHER_EXE_BASENAME} をサマリ Excel 同階層 UNC へ自動再配備する。
 * <p>同梱 {@link AppPaths#RDP_LAUNCHER_VERSION_BASENAME} の版数はリポジトリ直下 {@code version.txt} と同一。
 * 正本は {@code version.txt}。{@code scripts/build-rdp-remote-launcher.ps1} および pre-commit で同期する。
 * <p>版数だけが先行して上がった場合（exe 未再ビルド）は、共有先 exe の SHA-256 と同梱 exe を照合し転送を抑止する。
 */
public final class RdpRemoteLauncherDeployer {

    private static final CopyOption[] COPY_OPTIONS =
            new CopyOption[] {StandardCopyOption.REPLACE_EXISTING};

    private static final String BUNDLED_RESOURCE_PREFIX = "/jp/co/pm/ai/desktop/rdp-launcher/";

    private static final int HASH_BUFFER_SIZE = 8192;

    private RdpRemoteLauncherDeployer() {}

    public record DeployOutcome(boolean copied, boolean upToDate, Optional<String> message) {

        /** 配備成功（最新済み・コピー完了・自動配備オフ）または転送不要。 */
        public boolean succeeded() {
            return copied || upToDate;
        }
    }

    public static boolean isAutoDeployEnabled(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String raw = u.get(AppPaths.KEY_PM_AI_RDP_LAUNCHER_AUTO_DEPLOY);
        if (raw == null || raw.isBlank()) {
            return true;
        }
        String v = raw.trim().toLowerCase(Locale.ROOT);
        return !("0".equals(v) || "false".equals(v) || "off".equals(v) || "no".equals(v));
    }

    public static DeployOutcome ensureDeployed(Map<String, String> ui) {
        return ensureDeployed(ui, null);
    }

    public static DeployOutcome ensureDeployed(Map<String, String> ui, Consumer<String> log) {
        Map<String, String> env = ui != null ? ui : Map.of();
        if (!isAutoDeployEnabled(env)) {
            return new DeployOutcome(false, true, Optional.of("ランチャー自動再配備は無効です。"));
        }
        return deployIfNeeded(env, log, false);
    }

    /** 版比較を行わず同梱 exe / version.txt を共有先へ上書きコピーする。 */
    public static DeployOutcome forceDeploy(Map<String, String> ui) {
        return forceDeploy(ui, null);
    }

    public static DeployOutcome forceDeploy(Map<String, String> ui, Consumer<String> log) {
        Map<String, String> env = ui != null ? ui : Map.of();
        return deployIfNeeded(env, log, true);
    }

    /** 同梱 exe の実体コピーが必要か（版数のみの差分では false）。強制転送は含まない。 */
    public static boolean needsExeDeploy(Map<String, String> ui) {
        Map<String, String> env = ui != null ? ui : Map.of();
        if (!isAutoDeployEnabled(env)) {
            return false;
        }
        return needsExeCopy(env, false);
    }

    private static DeployOutcome deployIfNeeded(
            Map<String, String> env, Consumer<String> log, boolean force) {
        Path deployExe = AppPaths.resolveRdpLauncherExe(env);
        Path deployVersion = AppPaths.resolveRdpLauncherVersionFile(env);
        Path deployDir = deployExe.getParent();

        Optional<BigDecimal> bundledVer = readBundledVersion();
        if (bundledVer.isEmpty()) {
            String msg = "同梱ランチャーの版情報が見つかりません。";
            if (log != null) {
                log.accept(msg);
            }
            return new DeployOutcome(false, false, Optional.of(msg));
        }

        Optional<BigDecimal> sharedVer = parseVersionFile(deployVersion);
        boolean missingExe = !Files.isRegularFile(deployExe);
        boolean needsCopy = force || needsExeCopy(env, force);

        if (!needsCopy) {
            String msg;
            if (needsVersionSync(bundledVer, sharedVer)) {
                try {
                    if (deployDir != null) {
                        Files.createDirectories(deployDir);
                    }
                    copyBundledResource(AppPaths.RDP_LAUNCHER_VERSION_BASENAME, deployVersion);
                    msg =
                            "ランチャーは最新です（版数 "
                                    + formatVersion(bundledVer)
                                    + " を同期しました）。";
                } catch (IOException ex) {
                    msg = "ランチャーは最新です（" + formatVersion(sharedVer) + "）。";
                }
            } else {
                msg = "ランチャーは最新です（" + formatVersion(sharedVer) + "）。";
            }
            if (log != null) {
                log.accept(msg);
            }
            return new DeployOutcome(false, true, Optional.of(msg));
        }

        try {
            copyLauncherFiles(deployExe, deployVersion, deployDir);
            String msg =
                    force
                            ? "ランチャーを強制転送しました（"
                                    + formatVersion(sharedVer)
                                    + " → "
                                    + formatVersion(bundledVer)
                                    + "）。"
                            : missingExe
                                    ? "ランチャーを配備しました（" + formatVersion(bundledVer) + "）。"
                                    : "ランチャーを再配備しました（"
                                            + formatVersion(sharedVer)
                                            + " → "
                                            + formatVersion(bundledVer)
                                            + "）。";
            if (log != null) {
                log.accept(msg);
            }
            return new DeployOutcome(true, false, Optional.of(msg));
        } catch (IOException ex) {
            String msg = formatDeployFailure(ex);
            if (log != null) {
                log.accept(msg);
            }
            return new DeployOutcome(false, false, Optional.of(msg));
        }
    }

    private static String formatDeployFailure(IOException ex) {
        String detail = ex.getMessage() != null ? ex.getMessage() : ex.toString();
        if (looksLikeFileInUse(ex)) {
            return "ランチャー配備に失敗しました（接続先で "
                    + AppPaths.RDP_LAUNCHER_EXE_BASENAME
                    + " が使用中の可能性があります。"
                    + " リモートデスクトップの既存セッションを終了してから再転送してください）: "
                    + detail;
        }
        return "ランチャー配備に失敗しました: " + detail;
    }

    static boolean looksLikeFileInUse(IOException ex) {
        if (ex == null) {
            return false;
        }
        String msg = ex.getMessage();
        if (msg == null || msg.isBlank()) {
            return false;
        }
        String lower = msg.toLowerCase(Locale.ROOT);
        return lower.contains("being used by another process")
                || lower.contains("used by another process")
                || lower.contains("process cannot access")
                || lower.contains("sharing violation")
                || msg.contains("使用中")
                || msg.contains("別のプロセス")
                || msg.contains("共有違反");
    }

    private static boolean needsExeCopy(Map<String, String> env, boolean force) {
        if (force) {
            return true;
        }
        Path deployExe = AppPaths.resolveRdpLauncherExe(env);
        if (!Files.isRegularFile(deployExe)) {
            return true;
        }
        Optional<BigDecimal> bundledVer = readBundledVersion();
        if (bundledVer.isEmpty()) {
            return false;
        }
        Optional<String> bundledHash =
                sha256HexBundledResource(AppPaths.RDP_LAUNCHER_EXE_BASENAME);
        Optional<String> sharedHash = sha256HexFile(deployExe);
        if (bundledHash.isPresent()
                && sharedHash.isPresent()
                && bundledHash.get().equals(sharedHash.get())) {
            return false;
        }
        return true;
    }

    private static boolean needsVersionSync(
            Optional<BigDecimal> bundledVer, Optional<BigDecimal> sharedVer) {
        if (bundledVer.isEmpty()) {
            return false;
        }
        return sharedVer.isEmpty() || bundledVer.get().compareTo(sharedVer.get()) != 0;
    }

    static Optional<String> sha256HexFile(Path file) {
        if (!Files.isRegularFile(file)) {
            return Optional.empty();
        }
        try (InputStream in = Files.newInputStream(file)) {
            return sha256Hex(in);
        } catch (Exception ex) {
            return Optional.empty();
        }
    }

    static Optional<String> sha256HexBundledResource(String basename) {
        String resourcePath = BUNDLED_RESOURCE_PREFIX + basename;
        try (InputStream in = RdpRemoteLauncherDeployer.class.getResourceAsStream(resourcePath)) {
            if (in == null) {
                return Optional.empty();
            }
            return sha256Hex(in);
        } catch (Exception ex) {
            return Optional.empty();
        }
    }

    private static Optional<String> sha256Hex(InputStream in) throws Exception {
        MessageDigest digest = MessageDigest.getInstance("SHA-256");
        byte[] buffer = new byte[HASH_BUFFER_SIZE];
        int read;
        while ((read = in.read(buffer)) > 0) {
            digest.update(buffer, 0, read);
        }
        return Optional.of(HexFormat.of().formatHex(digest.digest()));
    }

    private static void copyLauncherFiles(Path deployExe, Path deployVersion, Path deployDir)
            throws IOException {
        if (deployDir != null) {
            Files.createDirectories(deployDir);
        }
        copyBundledResource(AppPaths.RDP_LAUNCHER_EXE_BASENAME, deployExe);
        copyBundledResource(AppPaths.RDP_LAUNCHER_VERSION_BASENAME, deployVersion);
    }

    private static void copyBundledResource(String basename, Path target) throws IOException {
        String resourcePath = BUNDLED_RESOURCE_PREFIX + basename;
        try (InputStream in =
                RdpRemoteLauncherDeployer.class.getResourceAsStream(resourcePath)) {
            if (in == null) {
                throw new IOException("同梱リソースが見つかりません: " + resourcePath);
            }
            Path parent = target.getParent();
            if (parent != null) {
                Files.createDirectories(parent);
            }
            Files.copy(in, target, COPY_OPTIONS);
        }
    }

    static Optional<BigDecimal> readBundledVersion() {
        String resourcePath = BUNDLED_RESOURCE_PREFIX + AppPaths.RDP_LAUNCHER_VERSION_BASENAME;
        try (InputStream in = RdpRemoteLauncherDeployer.class.getResourceAsStream(resourcePath)) {
            if (in == null) {
                return Optional.empty();
            }
            String raw = new String(in.readAllBytes(), StandardCharsets.UTF_8).trim();
            if (raw.isEmpty()) {
                return Optional.empty();
            }
            String firstLine = raw.lines().findFirst().orElse("").trim();
            if (firstLine.isEmpty()) {
                return Optional.empty();
            }
            return Optional.of(new BigDecimal(firstLine));
        } catch (Exception ex) {
            return Optional.empty();
        }
    }

    static Optional<BigDecimal> parseVersionFile(Path file) {
        if (!Files.isRegularFile(file)) {
            return Optional.empty();
        }
        try {
            String raw = Files.readString(file, StandardCharsets.UTF_8).trim();
            if (raw.isEmpty()) {
                return Optional.empty();
            }
            String firstLine = raw.lines().findFirst().orElse("").trim();
            if (firstLine.isEmpty()) {
                return Optional.empty();
            }
            return Optional.of(new BigDecimal(firstLine));
        } catch (Exception ex) {
            return Optional.empty();
        }
    }

    private static String formatVersion(Optional<BigDecimal> version) {
        return version.map(BigDecimal::toPlainString).orElse("未配備");
    }
}
