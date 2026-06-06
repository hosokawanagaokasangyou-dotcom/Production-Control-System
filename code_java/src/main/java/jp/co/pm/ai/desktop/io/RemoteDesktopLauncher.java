package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Optional;

/**
 * Windows リモートデスクトップ接続（{@code mstsc.exe}）を .rdp プロファイルで起動する。
 */
public final class RemoteDesktopLauncher {

    private RemoteDesktopLauncher() {}

    public static boolean isSupportedPlatform() {
        return isWindows();
    }

    /**
     * @param rdpProfile .rdp ファイルの絶対パス
     * @throws IOException 起動失敗・未対応プラットフォーム・ファイル不正
     */
    public static void launch(Path rdpProfile) throws IOException {
        Path abs = validateRdpProfile(rdpProfile);
        if (!isSupportedPlatform()) {
            throw new IOException("リモートデスクトップの起動は Windows のみ対応です。");
        }
        Path mstsc = resolveMstscExe();
        if (mstsc == null) {
            throw new IOException("mstsc.exe が見つかりません。");
        }
        startDetached(List.of(mstsc.toString(), abs.toString()));
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

    private static void startDetached(List<String> command) throws IOException {
        ProcessBuilder pb = new ProcessBuilder(command);
        pb.redirectErrorStream(true);
        pb.start();
    }

    private static boolean isWindows() {
        return System.getProperty("os.name", "").toLowerCase(Locale.ROOT).contains("windows");
    }
}
