package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.TreeMap;

/**
 * 共有 UNC 上の {@code RAP設定.ini}（接続先 RDP ランチャー向け）の読み書き。
 */
public final class RdpRemoteLauncherIni {

    public static final String SELECTED_SLOT_KEY = "起動プログラム番号";
    public static final int MAX_SLOTS = 9;

    /** exe パスと引数。 */
    public record Command(String executable, String arguments) {}

    private int selectedSlot = 1;
    private final Map<Integer, Command> slots = new TreeMap<>();

    public int selectedSlot() {
        return selectedSlot;
    }

    public void setSelectedSlot(int slot) {
        if (slot < 1 || slot > MAX_SLOTS) {
            throw new IllegalArgumentException("起動プログラム番号は 1～" + MAX_SLOTS + " です: " + slot);
        }
        selectedSlot = slot;
    }

    public Command getSlotCommand(int slot) {
        return slots.getOrDefault(slot, new Command("", ""));
    }

    /** ini 1 行分（{@code "exe" [引数...]}）の文字列。 */
    public String getSlot(int slot) {
        Command command = slots.get(slot);
        if (command == null || command.executable().isBlank()) {
            return "";
        }
        return formatSlotIniValue(command.executable(), command.arguments());
    }

    public void setSlotCommand(int slot, String program, String arguments) {
        if (slot < 1 || slot > MAX_SLOTS) {
            throw new IllegalArgumentException("スロット番号は 1～" + MAX_SLOTS + " です: " + slot);
        }
        String programTrimmed = program != null ? program.trim() : "";
        String argsTrimmed = arguments != null ? arguments.trim() : "";
        if (programTrimmed.isEmpty()) {
            slots.remove(slot);
        } else {
            slots.put(slot, new Command(programTrimmed, argsTrimmed));
        }
    }

    /** @deprecated UI からは {@link #setSlotCommand(int, String, String)} を使用 */
    @Deprecated
    public void setSlot(int slot, String commandLine) {
        if (commandLine == null || commandLine.isBlank()) {
            slots.remove(slot);
            return;
        }
        Command parsed = parseCommandLine(commandLine);
        setSlotCommand(slot, parsed.executable(), parsed.arguments());
    }

    public static RdpRemoteLauncherIni load(Path path) throws IOException {
        Objects.requireNonNull(path, "path");
        RdpRemoteLauncherIni ini = new RdpRemoteLauncherIni();
        if (!Files.isRegularFile(path)) {
            return ini;
        }
        List<String> lines = Files.readAllLines(path, StandardCharsets.UTF_8);
        for (String rawLine : lines) {
            String line = rawLine.trim();
            if (line.isEmpty() || line.startsWith("#") || line.startsWith(";")) {
                continue;
            }
            int eq = line.indexOf('=');
            if (eq <= 0) {
                continue;
            }
            String key = line.substring(0, eq).trim();
            String value = line.substring(eq + 1).trim();
            if (SELECTED_SLOT_KEY.equals(key)) {
                try {
                    int slot = Integer.parseInt(value);
                    if (slot >= 1 && slot <= MAX_SLOTS) {
                        ini.selectedSlot = slot;
                    }
                } catch (NumberFormatException ignored) {
                    // keep default
                }
                continue;
            }
            try {
                int slot = Integer.parseInt(key);
                if (slot >= 1 && slot <= MAX_SLOTS && !value.isEmpty()) {
                    Command parsed = parseCommandLine(value);
                    ini.setSlotCommand(slot, parsed.executable(), parsed.arguments());
                }
            } catch (RuntimeException ignored) {
                // ignore unknown or malformed keys
            }
        }
        return ini;
    }

    public void save(Path path) throws IOException {
        Objects.requireNonNull(path, "path");
        Path parent = path.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        List<String> lines = new ArrayList<>();
        lines.add(SELECTED_SLOT_KEY + "=" + selectedSlot);
        for (Map.Entry<Integer, Command> entry : slots.entrySet()) {
            Command command = entry.getValue();
            if (command != null && !command.executable().isBlank()) {
                lines.add(
                        entry.getKey()
                                + "="
                                + formatSlotIniValue(command.executable(), command.arguments()));
            }
        }
        Files.write(path, lines, StandardCharsets.UTF_8);
    }

    /**
     * ini スロット行の値を生成する。exe は常に {@code "..."} で囲む（パス空白対策）。
     */
    public static String formatSlotIniValue(String program, String arguments) {
        if (program == null || program.isBlank()) {
            throw new IllegalArgumentException("プログラムパスが空です。");
        }
        String quoted = "\"" + program.trim().replace("\"", "\"\"") + "\"";
        String args = arguments != null ? arguments.trim() : "";
        if (args.isEmpty()) {
            return quoted;
        }
        return quoted + " " + args;
    }

    /**
     * 1 行の「"exe" [引数...]」または「exe [引数...]」を分割する。
     */
    public static Command parseCommandLine(String line) {
        if (line == null || line.isBlank()) {
            throw new IllegalArgumentException("コマンド行が空です。");
        }
        String trimmed = line.trim();
        if (trimmed.startsWith("\"")) {
            int end = findClosingQuote(trimmed, 1);
            if (end < 0) {
                throw new IllegalArgumentException("引用符が閉じられていません: " + line);
            }
            String executable = trimmed.substring(1, end).replace("\"\"", "\"");
            String arguments =
                    end + 1 < trimmed.length() ? trimmed.substring(end + 1).trim() : "";
            return new Command(executable, arguments);
        }
        int space = trimmed.indexOf(' ');
        if (space < 0) {
            return new Command(trimmed, "");
        }
        return new Command(trimmed.substring(0, space), trimmed.substring(space + 1).trim());
    }

    private static int findClosingQuote(String text, int fromIndex) {
        for (int i = fromIndex; i < text.length(); i++) {
            if (text.charAt(i) != '"') {
                continue;
            }
            if (i + 1 < text.length() && text.charAt(i + 1) == '"') {
                i++;
                continue;
            }
            return i;
        }
        return -1;
    }

    public String validateMessageForSave() {
        if (selectedSlot < 1 || selectedSlot > MAX_SLOTS) {
            return "起動プログラム番号は 1～" + MAX_SLOTS + " を指定してください。";
        }
        Command selected = getSlotCommand(selectedSlot);
        if (selected.executable().isBlank()) {
            return "起動プログラム番号 " + selectedSlot + " のプログラムパスが空です。";
        }
        for (Map.Entry<Integer, Command> entry : slots.entrySet()) {
            Command command = entry.getValue();
            if (command == null || command.executable().isBlank()) {
                continue;
            }
            try {
                formatSlotIniValue(command.executable(), command.arguments());
            } catch (IllegalArgumentException ex) {
                return "スロット " + entry.getKey() + " が不正です: " + ex.getMessage();
            }
        }
        return null;
    }

    public int highestDefinedSlot() {
        if (slots.isEmpty()) {
            return 0;
        }
        int max = 0;
        for (Integer key : slots.keySet()) {
            if (key > max) {
                max = key;
            }
        }
        return max;
    }

    /** UI 向け: 1..max(3, highest) のスロット行数。 */
    public int visibleSlotCount() {
        int highest = highestDefinedSlot();
        return Math.min(MAX_SLOTS, Math.max(3, highest == 0 ? 3 : highest));
    }
}
