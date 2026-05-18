package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import java.util.function.Consumer;

/**
 * 実行中の {@code PMD.exe} がロックするファイルを、終了後に PowerShell スクリプトで適用する。
 */
public final class PortableBundleUpdateLauncher {

  public static final String UPDATE_SCRIPT_NAME = "pmd-apply-portable-update.ps1";

  private PortableBundleUpdateLauncher() {}

  /**
   * インストール直下または {@code pm-ai-data/scripts/} に同梱された更新スクリプトを起動する。
   *
   * @return 起動に成功したら true
   */
  public static boolean launchDeferredDesktopApply(
      Path installRoot,
      Path stagingRoot,
      long waitPid,
      String versionLabel,
      Path canonicalPath,
      Consumer<String> log)
      throws IOException {
    Objects.requireNonNull(installRoot, "installRoot");
    Objects.requireNonNull(stagingRoot, "stagingRoot");
    Path script = resolveUpdateScript(installRoot);
    if (script == null || !Files.isRegularFile(script)) {
      throw new IOException(
              "更新スクリプトが見つかりません: "
                      + UPDATE_SCRIPT_NAME
                      + "（インストール先または pm-ai-data/scripts/）");
    }
    Path logFile =
            PortableBundlePendingUpdate.userStateDirectory()
                    .resolve("pending-portable-update-apply.log");
    Files.createDirectories(logFile.getParent());

    List<String> cmd = new ArrayList<>();
    String os = System.getProperty("os.name", "").toLowerCase(Locale.ROOT);
    if (os.contains("win")) {
      cmd.add("powershell.exe");
      cmd.add("-NoProfile");
      cmd.add("-ExecutionPolicy");
      cmd.add("Bypass");
      cmd.add("-File");
      cmd.add(script.toAbsolutePath().normalize().toString());
    } else {
      throw new IOException("デスクトップ本体の終了後更新は Windows 専用です。");
    }
    cmd.add("-InstallRoot");
    cmd.add(installRoot.toAbsolutePath().normalize().toString());
    cmd.add("-StagingRoot");
    cmd.add(stagingRoot.toAbsolutePath().normalize().toString());
    cmd.add("-WaitPid");
    cmd.add(Long.toString(waitPid));
    cmd.add("-LogFile");
    cmd.add(logFile.toAbsolutePath().normalize().toString());
    if (versionLabel != null && !versionLabel.isBlank()) {
      cmd.add("-VersionLabel");
      cmd.add(versionLabel);
    }
    if (canonicalPath != null) {
      cmd.add("-CanonicalPath");
      cmd.add(canonicalPath.toAbsolutePath().normalize().toString());
    }

    ProcessBuilder pb = new ProcessBuilder(cmd);
    pb.inheritIO();
    pb.start();
    if (log != null) {
      log.accept(
              PortableBundleSelfUpdater.PORTABLE_SYNC_LOG_PREFIX
                      + "終了後更新スクリプトを起動: "
                      + PortableBundleSelfUpdater.safePathForLog(script));
      log.accept(
              PortableBundleSelfUpdater.PORTABLE_SYNC_LOG_PREFIX
                      + "適用ログ: "
                      + PortableBundleSelfUpdater.safePathForLog(logFile));
    }
    return true;
  }

  /** {@code installRoot} 直下、次に {@code pm-ai-data/scripts/}。 */
  public static Path resolveUpdateScript(Path installRoot) {
    if (installRoot == null) {
      return null;
    }
    Path atRoot = installRoot.resolve(UPDATE_SCRIPT_NAME);
    if (Files.isRegularFile(atRoot)) {
      return atRoot;
    }
    Path underData = installRoot.resolve("pm-ai-data").resolve("scripts").resolve(UPDATE_SCRIPT_NAME);
    if (Files.isRegularFile(underData)) {
      return underData;
    }
    return null;
  }

  /**
   * 前回の終了後更新が未完了のとき、コールドスタートでステージングを適用する。
   *
   * @return 適用したら true
   */
  public static boolean tryApplyStagedBundleOnColdStart(Path installRoot, Consumer<String> log) {
    Objects.requireNonNull(installRoot, "installRoot");
    Optional<PortableBundlePendingUpdate> pending = PortableBundlePendingUpdate.readIfPresent();
    if (pending.isEmpty()) {
      return false;
    }
    PortableBundlePendingUpdate p = pending.get();
    Path expectedInstall = installRoot.toAbsolutePath().normalize();
    if (!p.installRootPath().equals(expectedInstall)) {
      if (log != null) {
        log.accept(
                PortableBundleSelfUpdater.PORTABLE_SYNC_LOG_PREFIX
                        + "保留中の更新は別インストール向けのためスキップ: "
                        + p.installRoot());
      }
      return false;
    }
    Path staging = p.stagingRootPath();
    if (!PortableBundleSelfUpdater.hasDesktopInstallLayout(staging)) {
      if (log != null) {
        log.accept(
                PortableBundleSelfUpdater.PORTABLE_SYNC_LOG_PREFIX
                        + "保留ステージングが無効のため削除: "
                        + PortableBundleSelfUpdater.safePathForLog(staging));
      }
      PortableBundlePendingUpdate.clear();
      PortableBundlePendingUpdate.clearStaging(staging);
      return false;
    }
    try {
      if (log != null) {
        log.accept(
                PortableBundleSelfUpdater.PORTABLE_SYNC_LOG_PREFIX
                        + "コールドスタート: 保留中のデスクトップ本体を適用します…");
      }
      PortableBundleSelfUpdater.syncDesktopInstallFromBundleRoot(staging, expectedInstall, log, null);
      PortableBundlePendingUpdate.clear();
      PortableBundlePendingUpdate.clearStaging(staging);
      if (log != null) {
        log.accept(
                PortableBundleSelfUpdater.PORTABLE_SYNC_LOG_PREFIX
                        + "コールドスタート: デスクトップ本体の保留適用が完了しました。");
      }
      return true;
    } catch (IOException e) {
      if (log != null) {
        log.accept(
                PortableBundleSelfUpdater.PORTABLE_SYNC_LOG_PREFIX
                        + "コールドスタート: 保留適用に失敗: "
                        + e.getMessage());
      }
      return false;
    }
  }
}
