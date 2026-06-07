package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** .rdp プロファイルの表示設定・リモート起動プログラムを編集する（UTF-16 LE）。 */
public final class RdpProfileEditor {

    private static final Pattern INT_SETTING =
            Pattern.compile("^([a-z0-9 ]+):i:(-?\\d+)$", Pattern.CASE_INSENSITIVE);
    private static final Pattern STRING_SETTING =
            Pattern.compile("^([a-z0-9 ]+):s:(.*)$", Pattern.CASE_INSENSITIVE);

    private RdpProfileEditor() {}

    /** 接続先で起動するプログラム設定（.rdp から読み取り）。 */
    public record RemoteStartupSettings(String programPath, String arguments) {}

    /**
     * リモート デスクトップの解像度を設定する。
     *
     * @return 署名行を削除した場合 {@code true}
     */
    public static boolean applyDesktopSize(Path rdpProfile, int width, int height) throws IOException {
        if (width <= 0 || height <= 0) {
            throw new IllegalArgumentException("width/height must be positive");
        }
        Path abs = RemoteDesktopLauncher.validateRdpProfile(rdpProfile);
        List<String> lines = readLines(abs);
        boolean removedSignature = false;
        List<String> out = new ArrayList<>(lines.size() + 2);
        boolean hasWidth = false;
        boolean hasHeight = false;
        boolean hasScreenMode = false;
        for (String rawLine : lines) {
            String line = stripBom(rawLine);
            if (line.isEmpty()) {
                out.add(line);
                continue;
            }
            if (isSignatureLine(line)) {
                removedSignature = true;
                continue;
            }
            Matcher intMatcher = INT_SETTING.matcher(line);
            if (intMatcher.matches()) {
                String key = normalizeKey(intMatcher.group(1));
                if ("desktopwidth".equals(key)) {
                    out.add("desktopwidth:i:" + width);
                    hasWidth = true;
                    continue;
                }
                if ("desktopheight".equals(key)) {
                    out.add("desktopheight:i:" + height);
                    hasHeight = true;
                    continue;
                }
                if ("screen mode id".equals(key)) {
                    out.add("screen mode id:i:2");
                    hasScreenMode = true;
                    continue;
                }
            }
            out.add(line);
        }
        if (!hasWidth) {
            out.add("desktopwidth:i:" + width);
        }
        if (!hasHeight) {
            out.add("desktopheight:i:" + height);
        }
        if (!hasScreenMode) {
            out.add("screen mode id:i:2");
        }
        writeLines(abs, out);
        return removedSignature;
    }

    /**
     * 接続先サーバー上で起動するプログラムを .rdp に書き込む（RemoteApp 形式）。
     *
     * @param remoteProgramPath 接続先の exe パス。空ならリモート自動起動を無効化。
     * @param arguments {@code remoteapplicationcmdline} に渡す引数文字列
     * @return 署名行を削除した場合 {@code true}
     */
    public static boolean applyRemoteStartupProgram(
            Path rdpProfile, String remoteProgramPath, String arguments) throws IOException {
        Path abs = RemoteDesktopLauncher.validateRdpProfile(rdpProfile);
        String program = remoteProgramPath != null ? remoteProgramPath.trim() : "";
        String args = arguments != null ? arguments.trim() : "";
        boolean enable = !program.isEmpty();
        List<String> lines = readLines(abs);
        boolean removedSignature = false;
        List<String> out = new ArrayList<>(lines.size() + 4);
        boolean hasMode = false;
        boolean hasName = false;
        boolean hasProgram = false;
        boolean hasCmdline = false;
        boolean hasAlternateShell = false;
        boolean hasDisableCheck = false;
        for (String rawLine : lines) {
            String line = stripBom(rawLine);
            if (line.isEmpty()) {
                out.add(line);
                continue;
            }
            if (isSignatureLine(line)) {
                removedSignature = true;
                continue;
            }
            Matcher intMatcher = INT_SETTING.matcher(line);
            if (intMatcher.matches()) {
                String key = normalizeKey(intMatcher.group(1));
                if ("remoteapplicationmode".equals(key)) {
                    out.add("remoteapplicationmode:i:" + (enable ? 1 : 0));
                    hasMode = true;
                    continue;
                }
                if ("disableremoteappcheck".equals(key)) {
                    if (enable) {
                        out.add("disableremoteappcheck:i:1");
                    }
                    hasDisableCheck = true;
                    continue;
                }
            }
            Matcher strMatcher = STRING_SETTING.matcher(line);
            if (strMatcher.matches()) {
                String key = normalizeKey(strMatcher.group(1));
                if ("remoteapplicationprogram".equals(key)) {
                    out.add("remoteapplicationprogram:s:" + (enable ? program : ""));
                    hasProgram = true;
                    continue;
                }
                if ("remoteapplicationcmdline".equals(key)) {
                    out.add("remoteapplicationcmdline:s:" + (enable ? args : ""));
                    hasCmdline = true;
                    continue;
                }
                if ("remoteapplicationname".equals(key)) {
                    out.add("remoteapplicationname:s:" + (enable ? deriveRemoteAppName(program) : ""));
                    hasName = true;
                    continue;
                }
                if ("alternate shell".equals(key)) {
                    out.add("alternate shell:s:" + (enable ? "rdpinit.exe" : ""));
                    hasAlternateShell = true;
                    continue;
                }
            }
            out.add(line);
        }
        if (!hasMode) {
            out.add("remoteapplicationmode:i:" + (enable ? 1 : 0));
        }
        if (!hasProgram) {
            out.add("remoteapplicationprogram:s:" + (enable ? program : ""));
        }
        if (!hasCmdline) {
            out.add("remoteapplicationcmdline:s:" + (enable ? args : ""));
        }
        if (!hasName) {
            out.add("remoteapplicationname:s:" + (enable ? deriveRemoteAppName(program) : ""));
        }
        if (!hasAlternateShell) {
            out.add("alternate shell:s:" + (enable ? "rdpinit.exe" : ""));
        }
        if (enable && !hasDisableCheck) {
            out.add("disableremoteappcheck:i:1");
        }
        writeLines(abs, out);
        return removedSignature;
    }

    public static RemoteStartupSettings readRemoteStartupProgram(Path rdpProfile) throws IOException {
        Path abs = RemoteDesktopLauncher.validateRdpProfile(rdpProfile);
        String program = "";
        String args = "";
        for (String rawLine : readLines(abs)) {
            String line = stripBom(rawLine);
            Matcher strMatcher = STRING_SETTING.matcher(line);
            if (!strMatcher.matches()) {
                continue;
            }
            String key = normalizeKey(strMatcher.group(1));
            if ("remoteapplicationprogram".equals(key)) {
                program = strMatcher.group(2).trim();
            } else if ("remoteapplicationcmdline".equals(key)) {
                args = strMatcher.group(2).trim();
            }
        }
        return new RemoteStartupSettings(program, args);
    }

    static List<String> readLines(Path file) throws IOException {
        byte[] bytes = Files.readAllBytes(file);
        String text;
        if (bytes.length >= 2 && bytes[0] == (byte) 0xFF && bytes[1] == (byte) 0xFE) {
            text = new String(bytes, 2, bytes.length - 2, StandardCharsets.UTF_16LE);
        } else if (looksLikeUtf16LeRdp(bytes)) {
            text = new String(bytes, StandardCharsets.UTF_16LE);
        } else {
            text = Files.readString(file, StandardCharsets.UTF_8);
        }
        List<String> lines = new ArrayList<>();
        for (String part : text.split("\\R", -1)) {
            lines.add(part);
        }
        if (lines.isEmpty()) {
            lines.add("");
        }
        return lines;
    }

    static void writeLines(Path file, List<String> lines) throws IOException {
        String joined =
                String.join("\r\n", lines)
                        + (lines.isEmpty() || lines.get(lines.size() - 1).isEmpty() ? "" : "\r\n");
        byte[] body = joined.getBytes(StandardCharsets.UTF_16LE);
        byte[] out = new byte[body.length + 2];
        out[0] = (byte) 0xFF;
        out[1] = (byte) 0xFE;
        System.arraycopy(body, 0, out, 2, body.length);
        Files.write(file, out);
    }

    private static String deriveRemoteAppName(String program) {
        if (program == null || program.isBlank()) {
            return "RemoteApp";
        }
        String name = Path.of(program.replace('/', '\\')).getFileName().toString();
        int dot = name.lastIndexOf('.');
        return dot > 0 ? name.substring(0, dot) : name;
    }

    private static boolean isSignatureLine(String line) {
        return line.toLowerCase(Locale.ROOT).startsWith("signature:s:");
    }

    private static String normalizeKey(String key) {
        return key.strip().toLowerCase(Locale.ROOT);
    }

    private static String stripBom(String line) {
        if (line != null && !line.isEmpty() && line.charAt(0) == '\uFEFF') {
            return line.substring(1);
        }
        return line != null ? line : "";
    }

    private static boolean looksLikeUtf16LeRdp(byte[] bytes) {
        return containsUtf16LeMarker(bytes, "screen mode id:")
                || containsUtf16LeMarker(bytes, "full address:s:")
                || containsUtf16LeMarker(bytes, "desktopwidth:i:")
                || containsUtf16LeMarker(bytes, "desktopheight:i:")
                || containsUtf16LeMarker(bytes, "signature:s:")
                || containsUtf16LeMarker(bytes, "remoteapplicationmode:i:");
    }

    private static boolean containsUtf16LeMarker(byte[] bytes, String ascii) {
        byte[] pattern = new byte[ascii.length() * 2];
        for (int i = 0; i < ascii.length(); i++) {
            pattern[i * 2] = (byte) ascii.charAt(i);
        }
        if (bytes.length < pattern.length) {
            return false;
        }
        outer:
        for (int i = 0; i <= bytes.length - pattern.length; i++) {
            for (int j = 0; j < pattern.length; j++) {
                if (bytes[i + j] != pattern[j]) {
                    continue outer;
                }
            }
            return true;
        }
        return false;
    }
}
