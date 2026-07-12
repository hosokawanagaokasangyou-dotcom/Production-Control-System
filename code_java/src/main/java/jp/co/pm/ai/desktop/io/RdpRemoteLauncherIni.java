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

import jp.co.pm.ai.desktop.config.AladdinRpaLaunchArgs;
import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * 共有 UNC 上の {@code RPA設定.ini}（接続先 RDP ランチャー向け）の読み書き。
 */
public final class RdpRemoteLauncherIni {

    public static final String SELECTED_SLOT_KEY = "起動プログラム番号";
    /** 接続直前に Java が書くセッション操作者名（C# ランチャーが資格情報解決に使用）。 */
    public static final String OPERATOR_KEY = "操作者";
    /** 子プロセス終了後に RDP セッション操作を行うか（後方互換。{@link #SESSION_END_ACTION_KEY} が正）。 */
    public static final String DISCONNECT_ON_CHILD_EXIT_KEY = "終了時RDP切断";
    /** 子プロセス終了後のセッション操作（C# ランチャーが参照）。値: なし / 切断 / サインアウト */
    public static final String SESSION_END_ACTION_KEY = "終了時セッション操作";
    /**
     * 後方互換: 旧方式の接続時サインアウトフラグ（新方式はスロット {@link #SLOT_SIGN_OUT}={@link #SIGN_OUT_LAUNCHER_ARGS}）。
     */
    public static final String SIGN_OUT_ON_CONNECT_KEY = "接続時サインアウト";
    /** 接続先サインアウト専用の UI／ini 起動プロファイル番号。 */
    public static final int SLOT_SIGN_OUT = 99;
    /**
     * ini の {@link #SELECTED_SLOT_KEY}=0 … タスクスケジューラの RPA 二重起動抑止のみ（サインアウトしない）。
     */
    public static final int INI_SUPPRESS_SLOT = 0;
    /** {@link #INI_SUPPRESS_SLOT} の別名。 */
    public static final int INI_SIGN_OUT_SLOT = INI_SUPPRESS_SLOT;
    /** {@link #INI_SUPPRESS_SLOT} の別名（タスクスケジューラ抑止）。 */
    public static final int SLOT_DISABLED = INI_SUPPRESS_SLOT;
    public static final int MAX_SLOTS = 9;

    /** UI 起動時に用意する RPA プロファイル行数（{@link #SLOT_SIGN_OUT} 除く）。 */
    public static final int DEFAULT_INITIAL_RPA_PROFILE_ROWS = 5;

    /** タスクスケジューラ起動時に付与するサインアウト専用引数。 */
    public static final String SIGN_OUT_LAUNCHER_ARGS = "--signout";

    /** 起動プロファイル 99（{@link #SLOT_SIGN_OUT}）の表示名。 */
    public static final String SIGN_OUT_ONLY_PROFILE_NAME = "接続先サインアウトのみ";

    /** exe パスと引数。 */
    public record Command(String executable, String arguments) {}

    private int selectedSlot = 1;
    private boolean disconnectOnChildExit = true;
    private RdpSessionEndAction sessionEndAction = RdpSessionEndAction.SIGN_OUT;
    private boolean sessionEndActionExplicit;
    private String operatorName = "";
    private final Map<Integer, Command> slots = new TreeMap<>();

    public int selectedSlot() {
        return selectedSlot;
    }

    public void setSelectedSlot(int slot) {
        if (slot == INI_SUPPRESS_SLOT || slot == SLOT_SIGN_OUT) {
            selectedSlot = slot;
            return;
        }
        if (slot < 1 || slot > MAX_SLOTS) {
            throw new IllegalArgumentException("起動プログラム番号は 1～" + MAX_SLOTS + " です: " + slot);
        }
        selectedSlot = slot;
    }

    /** 起動プロファイル ComboBox 向け（99 は接続先サインアウトのみ）。 */
    public static boolean isSignOutOnlyProfile(int profileNumber) {
        return profileNumber == SLOT_SIGN_OUT;
    }

    /** ini の起動プログラム番号がタスクスケジューラ抑止専用か（0）。 */
    public static boolean isSuppressIniSlot(int iniSlot) {
        return iniSlot == INI_SUPPRESS_SLOT;
    }

    /** ini の起動プログラム番号が接続先サインアウト専用か（99）。 */
    public static boolean isSignOutIniSlot(int iniSlot) {
        return iniSlot == SLOT_SIGN_OUT;
    }

    /** @deprecated {@link #isSignOutIniSlot(int)} または {@link #isSuppressIniSlot(int)} を使用 */
    @Deprecated
    public static boolean isSignOutOnlyIniSlot(int iniSlot) {
        return isSignOutIniSlot(iniSlot);
    }

    public static String signOutOnlyProfileComboLabel() {
        return RdpLaunchProfile.signOutOnlyDefault().displayLabel();
    }

    public static String signOutOnlyProfileDetailText() {
        return "通常 mstsc で接続し、接続先タスクスケジューラが "
                + AppPaths.RDP_LAUNCHER_EXE_BASENAME
                + " 操作者名 を起動したとき、ini の "
                + SELECTED_SLOT_KEY
                + "="
                + SLOT_SIGN_OUT
                + " とスロット "
                + SLOT_SIGN_OUT
                + "="
                + SIGN_OUT_LAUNCHER_ARGS
                + " によりサインアウトします。alternate shell は使いません。"
                + " "
                + SELECTED_SLOT_KEY
                + "="
                + INI_SUPPRESS_SLOT
                + " だけではサインアウトしません。";
    }

    /** 保存・読込で UI プロファイル 99 を ini の起動プログラム番号 99 に対応させる。 */
    public void selectLaunchProfile(int profileNumber) {
        if (isSignOutOnlyProfile(profileNumber)) {
            selectedSlot = SLOT_SIGN_OUT;
            setSignOutSlotCommand();
            return;
        }
        setSelectedSlot(profileNumber);
    }

    public static boolean isSignOutSlotCommand(String executable) {
        return SIGN_OUT_LAUNCHER_ARGS.equalsIgnoreCase(
                executable != null ? executable.strip() : "");
    }

    /** スロット {@link #SLOT_SIGN_OUT} に {@link #SIGN_OUT_LAUNCHER_ARGS} を設定する。 */
    public void setSignOutSlotCommand() {
        slots.put(SLOT_SIGN_OUT, new Command(SIGN_OUT_LAUNCHER_ARGS, ""));
    }

    /** 子プロセス終了後にセッション操作を行う（後方互換）。 */
    public boolean disconnectOnChildExit() {
        return sessionEndAction.enabled();
    }

    public void setDisconnectOnChildExit(boolean disconnectOnChildExit) {
        sessionEndAction = disconnectOnChildExit ? RdpSessionEndAction.SIGN_OUT : RdpSessionEndAction.NONE;
        this.disconnectOnChildExit = disconnectOnChildExit;
    }

    public RdpSessionEndAction sessionEndAction() {
        return sessionEndAction;
    }

    public void setSessionEndAction(RdpSessionEndAction sessionEndAction) {
        this.sessionEndAction =
                sessionEndAction != null ? sessionEndAction : RdpSessionEndAction.SIGN_OUT;
        this.disconnectOnChildExit = this.sessionEndAction.enabled();
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
        if (slot == SLOT_SIGN_OUT) {
            if (!isSignOutSlotCommand(program)) {
                throw new IllegalArgumentException(
                        "スロット "
                                + SLOT_SIGN_OUT
                                + " には "
                                + SIGN_OUT_LAUNCHER_ARGS
                                + " のみ設定できます: "
                                + program);
            }
            setSignOutSlotCommand();
            return;
        }
        if (slot < 1 || slot > MAX_SLOTS) {
            throw new IllegalArgumentException("スロット番号は 1～" + MAX_SLOTS + " です: " + slot);
        }
        String programTrimmed =
                UncPathSegmentRepair.repair(
                        stripSurroundingQuotes(program != null ? program.trim() : ""));
        String argsTrimmed =
                arguments != null && !arguments.isBlank()
                        ? RpaScenarioArgumentSupport.repairScenarioArguments(arguments.trim())
                        : "";
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
        boolean sawSessionEndActionKey = false;
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
                    if (isSuppressIniSlot(slot)
                            || isSignOutIniSlot(slot)
                            || (slot >= 1 && slot <= MAX_SLOTS)) {
                        ini.selectedSlot = slot;
                    }
                } catch (NumberFormatException ignored) {
                    // keep default
                }
                continue;
            }
            if (DISCONNECT_ON_CHILD_EXIT_KEY.equals(key)) {
                ini.disconnectOnChildExit = parseBoolean(value, true);
                continue;
            }
            if (SESSION_END_ACTION_KEY.equals(key)) {
                ini.sessionEndAction =
                        RdpSessionEndAction.fromIniValue(value, RdpSessionEndAction.SIGN_OUT);
                ini.sessionEndActionExplicit = true;
                continue;
            }
            if (OPERATOR_KEY.equals(key)) {
                ini.operatorName = value != null ? value.strip() : "";
                continue;
            }
            try {
                int slot = Integer.parseInt(key);
                if (slot == SLOT_SIGN_OUT && !value.isEmpty()) {
                    Command parsed = parseCommandLine(value);
                    if (isSignOutSlotCommand(parsed.executable())) {
                        ini.setSignOutSlotCommand();
                    }
                    continue;
                }
                if (slot >= 1 && slot <= MAX_SLOTS && !value.isEmpty()) {
                    Command parsed = parseCommandLine(value);
                    ini.setSlotCommand(slot, parsed.executable(), parsed.arguments());
                }
            } catch (RuntimeException ignored) {
                // ignore unknown or malformed keys
            }
        }
        if (!ini.sessionEndActionExplicit) {
            ini.sessionEndAction =
                    ini.disconnectOnChildExit
                            ? RdpSessionEndAction.SIGN_OUT
                            : RdpSessionEndAction.NONE;
        }
        ini.disconnectOnChildExit = ini.sessionEndAction.enabled();
        return ini;
    }

    public void save(Path path) throws IOException {
        Objects.requireNonNull(path, "path");
        if (operatorName.isBlank()) {
            operatorName = readScalarValue(path, OPERATOR_KEY);
        }
        Path parent = path.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        Files.write(path, toIniLines(), StandardCharsets.UTF_8);
    }

    private List<String> toIniLines() {
        List<String> lines = new ArrayList<>();
        lines.add(SELECTED_SLOT_KEY + "=" + selectedSlot);
        lines.add(DISCONNECT_ON_CHILD_EXIT_KEY + "=" + (sessionEndAction.enabled() ? "1" : "0"));
        lines.add(SESSION_END_ACTION_KEY + "=" + sessionEndAction.iniValue());
        if (operatorName != null && !operatorName.isBlank()) {
            lines.add(OPERATOR_KEY + "=" + operatorName.strip());
        }
        for (Map.Entry<Integer, Command> entry : slots.entrySet()) {
            Command command = entry.getValue();
            if (command == null || command.executable().isBlank()) {
                continue;
            }
            if (entry.getKey() == SLOT_SIGN_OUT) {
                lines.add(SLOT_SIGN_OUT + "=" + SIGN_OUT_LAUNCHER_ARGS);
                continue;
            }
            lines.add(
                    entry.getKey()
                            + "="
                            + formatSlotIniValue(command.executable(), command.arguments()));
        }
        return lines;
    }

    /**
     * RDP 接続直前: 起動番号と指定スロットの RPA コマンド行を修復して ini に保存する。
     * タスクスケジューラ／{@link AppPaths#RDP_LAUNCHER_EXE_BASENAME} が参照する正本をここで揃える。
     */
    public static void writeLaunchContextBeforeConnect(
            Path path,
            int slot,
            String program,
            String arguments,
            RdpSessionEndAction sessionEndAction)
            throws IOException {
        Objects.requireNonNull(path, "path");
        if (slot < 1 || slot > MAX_SLOTS) {
            throw new IllegalArgumentException("起動プログラム番号は 1～" + MAX_SLOTS + " です: " + slot);
        }
        RdpRemoteLauncherIni ini = load(path);
        ini.setSelectedSlot(slot);
        ini.setSlotCommand(slot, program, arguments);
        if (sessionEndAction != null) {
            ini.setSessionEndAction(sessionEndAction);
        }
        ini.save(path);
    }

    /**
     * RDP 接続プロセス（mstsc）開始前に、タスクスケジューラが参照する起動番号を確定する。
     * 接続直後にタスクスケジューラが ini を読むため、{@link RemoteDesktopLauncher#launch} より先に呼ぶ。
     *
     * @deprecated 接続前は {@link #writeLaunchContextBeforeConnect} でスロット行も含めて保存すること。
     */
    @Deprecated
    public static void writeTaskSchedulerSlotBeforeConnect(Path path, int slot) throws IOException {
        restoreTaskSchedulerSlot(path, slot);
    }

    /** @see #writeTaskSchedulerSlotBeforeConnect(Path, int) */
    public static void writeTaskSchedulerSlotBeforeConnect(
            Path path, int slot, Map<String, String> ui) throws IOException {
        writeTaskSchedulerSlotBeforeConnect(path, slot);
    }

    /**
     * 接続直前に {@link #OPERATOR_KEY} 行のみ部分更新する（スロット定義は保持）。
     */
    public static void writeOperatorContext(Path path, String operatorName) throws IOException {
        mergeIniScalarKey(path, OPERATOR_KEY, operatorName != null ? operatorName.strip() : "");
    }

    private static String readScalarValue(Path path, String key) {
        if (path == null || key == null || key.isBlank() || !Files.isRegularFile(path)) {
            return "";
        }
        String prefix = key + "=";
        try {
            for (String rawLine : Files.readAllLines(path, StandardCharsets.UTF_8)) {
                String line = rawLine.trim();
                if (line.isEmpty() || line.startsWith("#") || line.startsWith(";")) {
                    continue;
                }
                if (line.startsWith(prefix)) {
                    return line.substring(prefix.length()).trim();
                }
            }
        } catch (IOException ignored) {
            // ignore
        }
        return "";
    }

    private static void mergeIniScalarKey(Path path, String key, String value) throws IOException {
        Objects.requireNonNull(path, "path");
        Objects.requireNonNull(key, "key");
        List<String> lines = new ArrayList<>();
        if (Files.isRegularFile(path)) {
            lines.addAll(Files.readAllLines(path, StandardCharsets.UTF_8));
        }
        String prefix = key + "=";
        boolean replaced = false;
        for (int i = 0; i < lines.size(); i++) {
            String trimmed = lines.get(i).trim();
            if (trimmed.isEmpty() || trimmed.startsWith("#") || trimmed.startsWith(";")) {
                continue;
            }
            if (!trimmed.startsWith(prefix)) {
                continue;
            }
            lines.set(i, key + "=" + value);
            replaced = true;
            break;
        }
        if (!replaced) {
            lines.add(key + "=" + value);
        }
        Path parent = path.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        Files.write(path, lines, StandardCharsets.UTF_8);
    }

    /**
     * ini の起動番号を {@link #INI_SUPPRESS_SLOT} にする（タスクスケジューラの RPA 二重起動抑止のみ）。
     */
    public static void writeTaskSchedulerSuppress(Path path) throws IOException {
        mergeIniScalarKey(path, SELECTED_SLOT_KEY, String.valueOf(INI_SUPPRESS_SLOT));
    }

    /** @see #writeTaskSchedulerSuppress(Path) */
    public static void writeTaskSchedulerSuppress(Path path, Map<String, String> ui)
            throws IOException {
        writeTaskSchedulerSuppress(path);
    }

    /**
     * プロファイル 99 用: {@link #SELECTED_SLOT_KEY}={@link #SLOT_SIGN_OUT} と
     * {@code 99=--signout} を書く。
     */
    public static void writeSignOutSlotRequest(Path path) throws IOException {
        RdpRemoteLauncherIni ini = load(path);
        ini.setSelectedSlot(SLOT_SIGN_OUT);
        ini.setSignOutSlotCommand();
        ini.save(path);
        clearSignOutOnConnectRequest(path);
    }

    /** @deprecated {@link #writeSignOutSlotRequest(Path)} を使用 */
    @Deprecated
    public static void writeSignOutOnConnectRequest(Path path) throws IOException {
        writeSignOutSlotRequest(path);
    }

    /** 旧方式フラグのクリア（後方互換）。 */
    public static void clearSignOutOnConnectRequest(Path path) throws IOException {
        mergeIniScalarKey(path, SIGN_OUT_ON_CONNECT_KEY, "0");
    }

    /**
     * RDP セッション終了後、タスクスケジューラ用に {@code RPA設定.ini} の起動番号を UI 値へ戻す。
     */
    public static void restoreTaskSchedulerSlot(Path path, int slot) throws IOException {
        Objects.requireNonNull(path, "path");
        if (slot < 1 || slot > MAX_SLOTS) {
            return;
        }
        mergeIniScalarKey(path, SELECTED_SLOT_KEY, String.valueOf(slot));
    }

    /** @see #restoreTaskSchedulerSlot(Path, int) */
    public static void restoreTaskSchedulerSlot(Path path, int slot, Map<String, String> ui)
            throws IOException {
        restoreTaskSchedulerSlot(path, slot);
    }

    /** デバッグ・UI 向け: ini の起動プログラム番号行（見つからなければ空文字）。 */
    public static String readIniHeadLine(Path path) {
        try {
            if (!Files.isRegularFile(path)) {
                return "";
            }
            for (String rawLine : Files.readAllLines(path, StandardCharsets.UTF_8)) {
                String line = rawLine.trim();
                if (line.isEmpty() || line.startsWith("#") || line.startsWith(";")) {
                    continue;
                }
                if (line.startsWith(SELECTED_SLOT_KEY + "=")) {
                    return line;
                }
            }
        } catch (IOException ignored) {
            // ignore
        }
        return "";
    }

    /**
     * ini スロット行の値を生成する。exe は常に {@code "..."} で囲む（パス空白対策）。
     * 引数に空白を含むトークンは {@code "..."} で囲む（cmd / CreateProcess 向け）。
     */
    public static String formatSlotIniValue(String program, String arguments) {
        if (program == null || program.isBlank()) {
            throw new IllegalArgumentException("プログラムパスが空です。");
        }
        String quoted = "\"" + program.trim().replace("\"", "\"\"") + "\"";
        String formattedArgs = formatArgumentsForProcess(arguments);
        if (formattedArgs.isEmpty()) {
            return quoted;
        }
        return quoted + " " + formattedArgs;
    }

    /** UI 表示用: 保存済み ini の引数から引用符を外して見せる。 */
    public static String argumentsForUiDisplay(String arguments) {
        if (arguments == null || arguments.isBlank()) {
            return "";
        }
        return String.join(" ", tokenizeArguments(arguments));
    }

    /** ini 引数に {@link AladdinRpaLaunchArgs#ETERNAL_FLAG} が含まれるか。 */
    public static boolean hasEternalFlag(String arguments) {
        return tokenizeArguments(arguments != null ? arguments : "").stream()
                .anyMatch(token -> AladdinRpaLaunchArgs.ETERNAL_FLAG.equalsIgnoreCase(token));
    }

    /** UI 向け: {@link AladdinRpaLaunchArgs#ETERNAL_FLAG} を除いた引数表示。 */
    public static String argumentsForUiDisplayWithoutEternal(String arguments) {
        return argumentsForUiDisplayWithoutManagedFlags(arguments);
    }

    /**
     * UI 向け: ランチャー付与フラグ（{@link AladdinRpaLaunchArgs#ID_FLAG} 等）と {@link AladdinRpaLaunchArgs#ETERNAL_FLAG}
     * を除いた表示。
     */
    public static String argumentsForUiDisplayWithoutManagedFlags(String arguments) {
        List<String> tokens = new ArrayList<>(tokenizeArguments(arguments != null ? arguments : ""));
        stripCredentialFlags(tokens);
        tokens.removeIf(token -> AladdinRpaLaunchArgs.ETERNAL_FLAG.equalsIgnoreCase(token));
        if (tokens.isEmpty()) {
            return "";
        }
        return String.join(" ", tokens);
    }

    /** UI／保存向け: {@code --scenario "path.ardrpa"} 形式の 1 シナリオ引数。 */
    public static String formatScenarioArgument(String scenarioPath) {
        if (scenarioPath == null || scenarioPath.isBlank()) {
            return "";
        }
        return formatTokensForIniArguments(
                List.of(
                        AladdinRpaLaunchArgs.SCENARIO_FLAG,
                        UncPathSegmentRepair.repair(scenarioPath.strip())));
    }

    /**
     * ini 保存向け: シナリオパスを {@link AladdinRpaLaunchArgs#SCENARIO_FLAG} 付きに正規化する。
     * 旧形式（.ardrpa パスのみ）も受け付ける。
     */
    public static String normalizeScenarioArguments(String arguments) {
        return RpaScenarioArgumentSupport.normalizeScenarioArguments(arguments);
    }

    /**
     * スロット保存用: シナリオ引数を正規化し {@link AladdinRpaLaunchArgs#ETERNAL_FLAG} を付与または除去する。
     */
    public static String mergeEternalFlag(String arguments, boolean eternal) {
        String normalized = normalizeScenarioArguments(arguments);
        List<String> tokens =
                normalized.isBlank()
                        ? new ArrayList<>()
                        : new ArrayList<>(tokenizeArguments(normalized));
        tokens.removeIf(token -> AladdinRpaLaunchArgs.ETERNAL_FLAG.equalsIgnoreCase(token));
        if (eternal) {
            tokens.add(AladdinRpaLaunchArgs.ETERNAL_FLAG);
        }
        if (tokens.isEmpty()) {
            return "";
        }
        return formatTokensForIniArguments(tokens);
    }

    private static void stripCredentialFlags(List<String> tokens) {
        for (int i = 0; i < tokens.size(); i++) {
            String token = tokens.get(i);
            if (!AladdinRpaLaunchArgs.ID_FLAG.equalsIgnoreCase(token)
                    && !AladdinRpaLaunchArgs.PASSWORD_FLAG.equalsIgnoreCase(token)) {
                continue;
            }
            tokens.remove(i);
            if (i < tokens.size()) {
                tokens.remove(i);
            }
            i = Math.max(-1, i - 1);
        }
    }

    private static void removeFlagWithValue(List<String> tokens, String flag) {
        for (int i = 0; i < tokens.size(); i++) {
            if (!flag.equalsIgnoreCase(tokens.get(i))) {
                continue;
            }
            tokens.remove(i);
            if (i < tokens.size()) {
                tokens.remove(i);
            }
            return;
        }
    }

    static String stripSurroundingQuotes(String value) {
        if (value == null || value.isBlank()) {
            return "";
        }
        String trimmed = value.strip();
        while (trimmed.length() >= 2 && trimmed.startsWith("\"") && trimmed.endsWith("\"")) {
            trimmed = trimmed.substring(1, trimmed.length() - 1).strip();
        }
        return trimmed.replace("\"\"", "\"");
    }

    private static String formatTokensForIniArguments(List<String> tokens) {
        StringBuilder out = new StringBuilder();
        for (String token : tokens) {
            if (token.isEmpty()) {
                continue;
            }
            if (out.length() > 0) {
                out.append(' ');
            }
            out.append(quoteArgumentIfNeeded(token));
        }
        return out.toString();
    }

    /**
     * 引数文字列をトークン化し、空白を含む各トークンを {@code "..."} で囲んで返す。
     */
    public static String formatArgumentsForProcess(String arguments) {
        if (arguments == null || arguments.isBlank()) {
            return "";
        }
        String trimmed = arguments.trim();
        if (!trimmed.startsWith("\"") && looksLikeSinglePathWithSpaces(trimmed)) {
            return quoteArgumentIfNeeded(trimmed);
        }
        List<String> tokens = tokenizeArguments(trimmed);
        StringBuilder out = new StringBuilder();
        for (String token : tokens) {
            if (token.isEmpty()) {
                continue;
            }
            if (out.length() > 0) {
                out.append(' ');
            }
            out.append(quoteArgumentIfNeeded(token));
        }
        return out.toString();
    }

    /** {@code \\server\...} や {@code Z:\...} の 1 パスに空白が含まれる入力（UI 未引用）向け。 */
    private static boolean looksLikeSinglePathWithSpaces(String value) {
        if (value.indexOf(' ') < 0) {
            return false;
        }
        if (value.startsWith("\\\\")) {
            return true;
        }
        return value.length() >= 2
                && value.charAt(1) == ':'
                && Character.isLetter(value.charAt(0));
    }

    static List<String> tokenizeArguments(String arguments) {
        List<String> tokens = new ArrayList<>();
        if (arguments == null || arguments.isBlank()) {
            return tokens;
        }
        StringBuilder current = new StringBuilder();
        boolean inQuotes = false;
        for (int i = 0; i < arguments.length(); i++) {
            char c = arguments.charAt(i);
            if (inQuotes) {
                if (c == '"') {
                    if (i + 1 < arguments.length() && arguments.charAt(i + 1) == '"') {
                        current.append('"');
                        i++;
                    } else {
                        inQuotes = false;
                    }
                } else {
                    current.append(c);
                }
            } else if (c == '"') {
                inQuotes = true;
            } else if (Character.isWhitespace(c)) {
                if (current.length() > 0) {
                    tokens.add(current.toString());
                    current.setLength(0);
                }
            } else {
                current.append(c);
            }
        }
        if (current.length() > 0) {
            tokens.add(current.toString());
        }
        return tokens;
    }

    private static String quoteArgumentIfNeeded(String token) {
        if (token.indexOf(' ') >= 0 || token.indexOf('\t') >= 0) {
            return "\"" + token.replace("\"", "\"\"") + "\"";
        }
        return token;
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
        if (isSuppressIniSlot(selectedSlot)) {
            return validateDefinedSlotCommands();
        }
        if (isSignOutIniSlot(selectedSlot)) {
            Command signOut = getSlotCommand(SLOT_SIGN_OUT);
            if (!isSignOutSlotCommand(signOut.executable())) {
                return "起動プログラム番号 "
                        + SLOT_SIGN_OUT
                        + " にはスロット "
                        + SLOT_SIGN_OUT
                        + "="
                        + SIGN_OUT_LAUNCHER_ARGS
                        + " が必要です。";
            }
            return validateDefinedSlotCommands();
        }
        if (selectedSlot < 1 || selectedSlot > MAX_SLOTS) {
            return "起動プログラム番号は 1～" + MAX_SLOTS + " を指定してください。";
        }
        Command selected = getSlotCommand(selectedSlot);
        if (selected.executable().isBlank()) {
            return "起動プログラム番号 " + selectedSlot + " のプログラムパスが空です。";
        }
        return validateDefinedSlotCommands();
    }

    private String validateDefinedSlotCommands() {
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

    /** UI 向け: 1..max({@link #DEFAULT_INITIAL_RPA_PROFILE_ROWS}, highest) のスロット行数。 */
    public int visibleSlotCount() {
        int highest = highestDefinedSlot();
        int floor = DEFAULT_INITIAL_RPA_PROFILE_ROWS;
        return Math.min(MAX_SLOTS, Math.max(floor, highest == 0 ? floor : highest));
    }

    /** 1～{@link #MAX_SLOTS} の RPA プロファイル番号の最大値（99 等は除外）。 */
    public static int maxRpaProfileNumber(Iterable<Integer> profileNumbers) {
        if (profileNumbers == null) {
            return 0;
        }
        int max = 0;
        for (Integer number : profileNumbers) {
            if (number != null && number >= 1 && number <= MAX_SLOTS && number > max) {
                max = number;
            }
        }
        return max;
    }

    private static boolean parseBoolean(String raw, boolean defaultValue) {
        if (raw == null || raw.isBlank()) {
            return defaultValue;
        }
        String v = raw.trim().toLowerCase(java.util.Locale.ROOT);
        return switch (v) {
            case "1", "true", "on", "yes" -> true;
            case "0", "false", "off", "no" -> false;
            default -> defaultValue;
        };
    }
}
