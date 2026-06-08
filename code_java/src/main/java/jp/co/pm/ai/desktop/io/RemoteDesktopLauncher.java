package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.OptionalLong;

/**
 * Windows リモートデスクトップ接続（{@code mstsc.exe}）を .rdp プロファイルで起動する。
 */
public final class RemoteDesktopLauncher {

    /** {@link #launch(Path, Map)} の結果。 */
    public record LaunchOutcome(
            Path rdpProfile,
            Optional<String> remoteStartupSummary,
            boolean signatureRemoved,
            /** 直接起動した mstsc の PID。 */
            OptionalLong mstscProcessId,
            /** セキュリティダイアログ自動操作時に mstsc PID を書き込むマーカーファイル。 */
            Optional<Path> mstscPidMarkerFile) {}

    private RemoteDesktopLauncher() {}

    public static boolean isSupportedPlatform() {
        return isWindows();
    }

    public static LaunchOutcome launch(Path rdpProfile) throws IOException {
        return launch(rdpProfile, Map.of());
    }

    /**
     * @param ui リモート起動プログラム設定の解決に使用
     */
    public static LaunchOutcome launch(Path rdpProfile, Map<String, String> ui) throws IOException {
        Path preferred = RdpFileSigner.resolvePreferredSignedProfilePath(rdpProfile, ui);
        Path abs = validateRdpProfile(preferred);
        if (!isSupportedPlatform()) {
            throw new IOException("リモートデスクトップの起動は Windows のみ対応です。");
        }
        Path mstsc = resolveMstscExe();
        if (mstsc == null) {
            throw new IOException("mstsc.exe が見つかりません。");
        }
        Map<String, String> env = ui != null ? ui : Map.of();
        String remoteProgram =
                RdpCompanionLauncher.resolveRemoteProgramPath(env).orElse("");
        String remoteArgs = RdpCompanionLauncher.resolveRemoteProgramArgs(env);
        boolean embedInProfile = RdpCompanionLauncher.isEmbedStartupInProfileEnabled(env);
        boolean signatureRemoved =
                embedInProfile
                        ? RdpProfileEditor.applyRemoteStartupProgram(abs, remoteProgram, remoteArgs)
                        : false;
        boolean displaySignatureRemoved = RdpLaunchDisplaySettings.applyToProfile(abs, env);
        signatureRemoved = signatureRemoved || displaySignatureRemoved;
        Optional<String> summary =
                embedInProfile ? RdpCompanionLauncher.formatEmbeddedSummary(env) : Optional.empty();
        OptionalLong mstscPidHint = OptionalLong.empty();
        Optional<Path> mstscPidMarker = Optional.empty();
        if (RdpSecurityDialogAutomator.isAutoConfirmEnabled(env)) {
            mstscPidMarker =
                    Optional.of(
                            RdpSecurityDialogAutomator.launchWithAutomatedConfirm(mstsc, abs, env));
        } else {
            Process started = startDetached(List.of(mstsc.toString(), abs.toString()));
            mstscPidHint = OptionalLong.of(started.pid());
        }
        OptionalLong mstscPid =
                mstscPidHint.isPresent() ? mstscPidHint : findLatestMstscPidForProfile(abs);
        return new LaunchOutcome(abs, summary, signatureRemoved, mstscPid, mstscPidMarker);
    }

    /** 起動直後の mstsc PID（自動確認 OFF 時は Process#pid、ON 時はスキャン）。 */
    private static OptionalLong findLatestMstscPidForProfile(Path rdpProfile) {
        return RdpMstscProcessFinder.scanForMstscProcessId(rdpProfile);
    }

    public static Optional<Path> resolveMstscExeOptional() {
        return Optional.ofNullable(resolveMstscExe());
    }

    public static Path validateRdpProfile(Path rdpProfile) throws IOException {
        if (rdpProfile == null) {
            throw new IOException("RDP プロファイルが未指定です。");
        }
        Path abs = rdpProfile.toAbsolutePath().normalize();
        if (!Files.isRegularFile(abs)) {
            throw new IOException("RDP プロファイルが見つかりません: " + abs);
        }
        String name = abs.getFileName().toString().toLowerCase(Locale.ROOT);
        if (!name.endsWith(".rdp")) {
            throw new IOException("RDP プロファイルは .rdp ファイルを指定してください: " + abs);
        }
        return abs;
    }

    private static Path resolveMstscExe() {
        List<Path> candidates = new ArrayList<>();
        String windir = System.getenv("SystemRoot");
        if (windir != null && !windir.isBlank()) {
            candidates.add(Path.of(windir.trim(), "System32", "mstsc.exe"));
        }
        candidates.add(Path.of("C:\\Windows\\System32\\mstsc.exe"));
        for (Path p : candidates) {
            if (Files.isRegularFile(p)) {
                return p.toAbsolutePath().normalize();
            }
        }
        return null;
    }

    private static Process startDetached(List<String> command) throws IOException {
        ProcessBuilder pb = new ProcessBuilder(command);
        pb.redirectErrorStream(true);
        return pb.start();
    }

    private static boolean isWindows() {
        return System.getProperty("os.name", "").toLowerCase(Locale.ROOT).contains("windows");
    }
}
