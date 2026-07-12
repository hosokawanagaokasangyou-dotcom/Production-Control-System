package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class RdpRemoteLauncherIniTest {

    @Test
    void parseCommandLine_notepadOnly() {
        RdpRemoteLauncherIni.Command cmd =
                RdpRemoteLauncherIni.parseCommandLine("C:\\Windows\\System32\\notepad.exe");
        assertEquals("C:\\Windows\\System32\\notepad.exe", cmd.executable());
        assertEquals("", cmd.arguments());
    }

    @Test
    void parseCommandLine_quotedExecutableOnly() {
        RdpRemoteLauncherIni.Command cmd =
                RdpRemoteLauncherIni.parseCommandLine("\"C:\\Windows\\System32\\notepad.exe\"");
        assertEquals("C:\\Windows\\System32\\notepad.exe", cmd.executable());
        assertEquals("", cmd.arguments());
    }

    @Test
    void parseCommandLine_quotedUncWithArgs() {
        String line =
                "\"\\\\192.168.0.101\\share\\002  加工G\\Aladdin_RPA_Studio.exe\" C:\\Users\\test\\file.ardrpa";
        RdpRemoteLauncherIni.Command cmd = RdpRemoteLauncherIni.parseCommandLine(line);
        assertEquals("\\\\192.168.0.101\\share\\002  加工G\\Aladdin_RPA_Studio.exe", cmd.executable());
        assertEquals("C:\\Users\\test\\file.ardrpa", cmd.arguments());
    }

    @Test
    void formatSlotIniValue_quotesProgramAlways() {
        assertEquals(
                "\"C:\\Windows\\System32\\notepad.exe\"",
                RdpRemoteLauncherIni.formatSlotIniValue("C:\\Windows\\System32\\notepad.exe", ""));
        assertEquals(
                "\"\\\\server\\share\\app.exe\" C:\\tmp\\a.ardrpa",
                RdpRemoteLauncherIni.formatSlotIniValue(
                        "\\\\server\\share\\app.exe", "C:\\tmp\\a.ardrpa"));
    }

    @Test
    void formatSlotIniValue_quotesArgumentsWithSpaces() {
        String arg =
                "\\\\192.168.0.101\\share\\scenarios\\工程別生産計画問い合わ せ.ardrpa";
        String formatted =
                RdpRemoteLauncherIni.formatSlotIniValue(
                        "Z:\\portable\\Aladdin_RPA_Studio.exe", arg);
        assertEquals(
                "\"Z:\\portable\\Aladdin_RPA_Studio.exe\" "
                        + "\"\\\\192.168.0.101\\share\\scenarios\\工程別生産計画問い合わ せ.ardrpa\"",
                formatted);
    }

    @Test
    void argumentsForUiDisplay_stripsQuotes() {
        assertEquals(
                "\\\\server\\a.ardrpa",
                RdpRemoteLauncherIni.argumentsForUiDisplay("\"\\\\server\\a.ardrpa\""));
        assertEquals(
                "\\\\server\\問い合わ せ.ardrpa",
                RdpRemoteLauncherIni.argumentsForUiDisplay(
                        "\"\\\\server\\問い合わ せ.ardrpa\""));
    }

    @Test
    void loadAndSave_persistsDisconnectOnChildExit(@TempDir Path tmp) throws Exception {
        Path iniPath = tmp.resolve("RAP設定.ini");
        Files.writeString(
                iniPath,
                """
                起動プログラム番号=1
                終了時RDP切断=0
                1="C:\\Windows\\System32\\notepad.exe"
                """,
                StandardCharsets.UTF_8);

        RdpRemoteLauncherIni loaded = RdpRemoteLauncherIni.load(iniPath);
        assertFalse(loaded.disconnectOnChildExit());
        assertEquals(RdpSessionEndAction.NONE, loaded.sessionEndAction());

        loaded.setDisconnectOnChildExit(true);
        Path out = tmp.resolve("out.ini");
        loaded.save(out);
        RdpRemoteLauncherIni again = RdpRemoteLauncherIni.load(out);
        assertTrue(again.disconnectOnChildExit());
        assertEquals(RdpSessionEndAction.SIGN_OUT, again.sessionEndAction());
        assertTrue(Files.readString(out, StandardCharsets.UTF_8).contains("終了時RDP切断=1"));
        assertTrue(
                Files.readString(out, StandardCharsets.UTF_8)
                        .contains(RdpRemoteLauncherIni.SESSION_END_ACTION_KEY + "=サインアウト"));
    }

    @Test
    void loadAndSave_persistsSessionEndActionDisconnect(@TempDir Path tmp) throws Exception {
        Path iniPath = tmp.resolve("RAP設定.ini");
        Files.writeString(
                iniPath,
                """
                起動プログラム番号=1
                終了時RDP切断=1
                終了時セッション操作=切断
                1="C:\\Windows\\System32\\notepad.exe"
                """,
                StandardCharsets.UTF_8);

        RdpRemoteLauncherIni loaded = RdpRemoteLauncherIni.load(iniPath);
        assertEquals(RdpSessionEndAction.DISCONNECT, loaded.sessionEndAction());

        loaded.setSessionEndAction(RdpSessionEndAction.SIGN_OUT);
        Path out = tmp.resolve("out.ini");
        loaded.save(out);
        RdpRemoteLauncherIni again = RdpRemoteLauncherIni.load(out);
        assertEquals(RdpSessionEndAction.SIGN_OUT, again.sessionEndAction());
    }

    @Test
    void loadAndSave_roundTripWithQuotes(@TempDir Path tmp) throws Exception {
        Path iniPath = tmp.resolve("RAP設定.ini");
        Files.writeString(
                iniPath,
                """
                起動プログラム番号=2
                1="C:\\Windows\\System32\\notepad.exe"
                2="\\\\server\\share\\002  加工G\\app.exe" arg1
                """,
                StandardCharsets.UTF_8);

        RdpRemoteLauncherIni loaded = RdpRemoteLauncherIni.load(iniPath);
        assertEquals(2, loaded.selectedSlot());
        assertEquals("C:\\Windows\\System32\\notepad.exe", loaded.getSlotCommand(1).executable());
        assertEquals("\\\\server\\share\\002  加工G\\app.exe", loaded.getSlotCommand(2).executable());
        assertEquals("arg1", loaded.getSlotCommand(2).arguments());

        loaded.setSlotCommand(2, "\\\\server\\share\\002  加工G\\app.exe", "arg1");
        Path out = tmp.resolve("out.ini");
        loaded.save(out);
        RdpRemoteLauncherIni again = RdpRemoteLauncherIni.load(out);
        assertEquals("arg1", again.getSlotCommand(2).arguments());
        assertTrue(again.getSlot(1).startsWith("\""));
        assertTrue(again.getSlot(1).endsWith("notepad.exe\""));
        assertTrue(Files.readString(out, StandardCharsets.UTF_8).contains("002  加工G"));
    }

    @Test
    void writeOperatorContext_mergesOperatorKey(@TempDir Path tmp) throws Exception {
        Path iniPath = tmp.resolve("RAP設定.ini");
        Files.writeString(
                iniPath,
                """
                起動プログラム番号=1
                1="Z:\\app.exe"
                """,
                StandardCharsets.UTF_8);

        RdpRemoteLauncherIni.writeOperatorContext(iniPath, "細川");

        String text = Files.readString(iniPath, StandardCharsets.UTF_8);
        assertTrue(text.contains(RdpRemoteLauncherIni.OPERATOR_KEY + "=細川"));
        RdpRemoteLauncherIni loaded = RdpRemoteLauncherIni.load(iniPath);
        assertEquals(1, loaded.selectedSlot());
    }

    @Test
    void writeTaskSchedulerSlotBeforeConnect_preservesOperator(@TempDir Path tmp) throws Exception {
        Path iniPath = tmp.resolve("RAP設定.ini");
        Files.writeString(
                iniPath,
                """
                起動プログラム番号=1
                操作者=古家
                2="Z:\\app.exe"
                """,
                StandardCharsets.UTF_8);

        RdpRemoteLauncherIni.writeTaskSchedulerSlotBeforeConnect(iniPath, 2);

        String text = Files.readString(iniPath, StandardCharsets.UTF_8);
        assertTrue(text.contains(RdpRemoteLauncherIni.OPERATOR_KEY + "=古家"));
    }

    @Test
    void writeTaskSchedulerSlotBeforeConnect_setsSpinnerSlot(@TempDir Path tmp) throws Exception {
        Path iniPath = tmp.resolve("RAP設定.ini");
        Files.writeString(
                iniPath,
                """
                起動プログラム番号=0
                2="Z:\\app.exe"
                """,
                StandardCharsets.UTF_8);

        RdpRemoteLauncherIni.writeTaskSchedulerSlotBeforeConnect(iniPath, 2);

        RdpRemoteLauncherIni loaded = RdpRemoteLauncherIni.load(iniPath);
        assertEquals(2, loaded.selectedSlot());
    }

    @Test
    void writeSignOutSlotRequest_setsSlot99AndSignOutCommand(@TempDir Path tmp) throws Exception {
        Path iniPath = tmp.resolve("細川_RPA設定.ini");
        Files.writeString(
                iniPath,
                """
                起動プログラム番号=2
                操作者=細川
                2="Z:\\app.exe"
                """,
                StandardCharsets.UTF_8);
        RdpRemoteLauncherIni.writeSignOutSlotRequest(iniPath);
        String text = Files.readString(iniPath, StandardCharsets.UTF_8);
        assertTrue(text.contains("起動プログラム番号=99"));
        assertTrue(text.contains("99=" + RdpRemoteLauncherIni.SIGN_OUT_LAUNCHER_ARGS));
        RdpRemoteLauncherIni loaded = RdpRemoteLauncherIni.load(iniPath);
        assertEquals(RdpRemoteLauncherIni.SLOT_SIGN_OUT, loaded.selectedSlot());
        assertTrue(RdpRemoteLauncherIni.isSignOutSlotCommand(loaded.getSlotCommand(99).executable()));
    }

    @Test
    void writeTaskSchedulerSuppress_setsSlotZeroPreservingSlots(@TempDir Path tmp) throws Exception {
        Path iniPath = tmp.resolve("RAP設定.ini");
        Files.writeString(
                iniPath,
                """
                起動プログラム番号=2
                終了時RDP切断=1
                2="Z:\\portable\\Aladdin_RPA_Studio.exe" "\\\\server\\a.ardrpa"
                """,
                StandardCharsets.UTF_8);

        RdpRemoteLauncherIni.writeTaskSchedulerSuppress(iniPath);

        RdpRemoteLauncherIni loaded = RdpRemoteLauncherIni.load(iniPath);
        assertEquals(RdpRemoteLauncherIni.INI_SUPPRESS_SLOT, loaded.selectedSlot());
        assertEquals("Z:\\portable\\Aladdin_RPA_Studio.exe", loaded.getSlotCommand(2).executable());
        assertTrue(Files.readString(iniPath, StandardCharsets.UTF_8).contains("起動プログラム番号=0"));
    }

    @Test
    void save_preservesOperatorKey(@TempDir Path tmp) throws Exception {
        Path iniPath = tmp.resolve("RAP設定.ini");
        Files.writeString(
                iniPath,
                """
                起動プログラム番号=1
                操作者=細川
                1="Z:\\app.exe"
                """,
                StandardCharsets.UTF_8);

        RdpRemoteLauncherIni ini = new RdpRemoteLauncherIni();
        ini.setSelectedSlot(1);
        ini.setSlotCommand(1, "Z:\\app.exe", "");
        ini.save(iniPath);

        String text = Files.readString(iniPath, StandardCharsets.UTF_8);
        assertTrue(text.contains(RdpRemoteLauncherIni.OPERATOR_KEY + "=細川"));
    }

    @Test
    void restoreTaskSchedulerSlot_restoresFromZero(@TempDir Path tmp) throws Exception {
        Path iniPath = tmp.resolve("RAP設定.ini");
        Files.writeString(
                iniPath,
                """
                起動プログラム番号=0
                2="Z:\\app.exe"
                """,
                StandardCharsets.UTF_8);

        RdpRemoteLauncherIni.restoreTaskSchedulerSlot(iniPath, 2, Map.of());

        RdpRemoteLauncherIni loaded = RdpRemoteLauncherIni.load(iniPath);
        assertEquals(2, loaded.selectedSlot());
        assertTrue(
                Files.readString(iniPath, StandardCharsets.UTF_8).contains("起動プログラム番号=2"));
    }

    @Test
    void load_parsesDisabledSlot(@TempDir Path tmp) throws Exception {
        Path iniPath = tmp.resolve("RAP設定.ini");
        Files.writeString(
                iniPath,
                """
                起動プログラム番号=0
                2="C:\\app.exe"
                """,
                StandardCharsets.UTF_8);

        RdpRemoteLauncherIni loaded = RdpRemoteLauncherIni.load(iniPath);
        assertEquals(RdpRemoteLauncherIni.INI_SUPPRESS_SLOT, loaded.selectedSlot());
    }

    @Test
    void load_parsesSignOutSlot99(@TempDir Path tmp) throws Exception {
        Path iniPath = tmp.resolve("RAP設定.ini");
        Files.writeString(
                iniPath,
                """
                起動プログラム番号=99
                99=--signout
                2="C:\\app.exe"
                """,
                StandardCharsets.UTF_8);

        RdpRemoteLauncherIni loaded = RdpRemoteLauncherIni.load(iniPath);
        assertEquals(RdpRemoteLauncherIni.SLOT_SIGN_OUT, loaded.selectedSlot());
        assertTrue(RdpRemoteLauncherIni.isSignOutSlotCommand(loaded.getSlotCommand(99).executable()));
    }

    @Test
    void validateMessageForSave_emptyProgramOnSelectedSlot() {
        RdpRemoteLauncherIni ini = new RdpRemoteLauncherIni();
        ini.setSelectedSlot(2);
        ini.setSlotCommand(1, "C:\\Windows\\System32\\notepad.exe", "");
        assertNotNull(ini.validateMessageForSave());
    }

    @Test
    void validateMessageForSave_ok() {
        RdpRemoteLauncherIni ini = new RdpRemoteLauncherIni();
        ini.setSlotCommand(1, "C:\\Windows\\System32\\notepad.exe", "");
        assertNull(ini.validateMessageForSave());
    }

    @Test
    void validateMessageForSave_signOutOnlyProfile_requiresSignOutSlot() {
        RdpRemoteLauncherIni ini = new RdpRemoteLauncherIni();
        ini.selectLaunchProfile(RdpRemoteLauncherIni.SLOT_SIGN_OUT);
        assertEquals(RdpRemoteLauncherIni.SLOT_SIGN_OUT, ini.selectedSlot());
        ini.setSlotCommand(1, "C:\\Windows\\System32\\notepad.exe", "");
        assertNull(ini.validateMessageForSave());
    }

    @Test
    void isSignOutOnlyProfile_recognizesSlot99() {
        assertTrue(RdpRemoteLauncherIni.isSignOutOnlyProfile(99));
        assertFalse(RdpRemoteLauncherIni.isSignOutOnlyProfile(1));
        assertTrue(RdpRemoteLauncherIni.isSuppressIniSlot(0));
        assertTrue(RdpRemoteLauncherIni.isSignOutIniSlot(99));
        assertEquals("99: 接続先サインアウトのみ", RdpRemoteLauncherIni.signOutOnlyProfileComboLabel());
    }

    @Test
    void setSlotCommand_acceptsSignOutSlot() {
        RdpRemoteLauncherIni ini = new RdpRemoteLauncherIni();
        ini.setSlotCommand(99, "--signout", "");
        assertTrue(RdpRemoteLauncherIni.isSignOutSlotCommand(ini.getSlotCommand(99).executable()));
    }

    @Test
    void setSlotCommand_rejectsNonSignOutOnSlot99() {
        RdpRemoteLauncherIni ini = new RdpRemoteLauncherIni();
        assertThrows(
                IllegalArgumentException.class,
                () -> ini.setSlotCommand(99, "M:\\launcher.exe", ""));
    }

    @Test
    void parseCommandLine_blankThrows() {
        assertThrows(IllegalArgumentException.class, () -> RdpRemoteLauncherIni.parseCommandLine("  "));
    }

    @Test
    void mergeEternalFlag_addsAndRemovesFlag() {
        assertEquals("--eternal", RdpRemoteLauncherIni.mergeEternalFlag("", true));
        String withScenario =
                RdpRemoteLauncherIni.mergeEternalFlag("\\\\server\\share\\a.ardrpa", true);
        assertTrue(withScenario.contains("--scenario"));
        assertTrue(withScenario.contains("a.ardrpa"));
        assertTrue(RdpRemoteLauncherIni.hasEternalFlag(withScenario));
        assertTrue(
                RdpRemoteLauncherIni.hasEternalFlag(
                        RdpRemoteLauncherIni.mergeEternalFlag(
                                "--scenario \\\\server\\share\\a.ardrpa --eternal", true)));
        assertFalse(
                RdpRemoteLauncherIni.hasEternalFlag(
                        RdpRemoteLauncherIni.mergeEternalFlag("\\\\server\\share\\a.ardrpa", false)));
    }

    @Test
    void normalizeScenarioArguments_preservesDoubleSpaceInUnquotedPath() {
        String path = "\\\\192.168.0.101\\share\\002  加工G\\a.ardrpa";
        assertEquals(
                "--scenario \"" + path + "\"",
                RdpRemoteLauncherIni.normalizeScenarioArguments("--scenario " + path));
        assertEquals(
                "--scenario \"" + path + "\"",
                RdpRemoteLauncherIni.normalizeScenarioArguments(
                        "--scenario \"" + path + "\""));
    }

    @Test
    void normalizeScenarioArguments_repairsCollapsed002Segment() {
        String wrong = "\\\\192.168.0.101\\share\\002 加工G\\a.ardrpa";
        String correct = "\\\\192.168.0.101\\share\\002  加工G\\a.ardrpa";
        assertEquals(
                "--scenario \"" + correct + "\"",
                RdpRemoteLauncherIni.normalizeScenarioArguments("--scenario " + wrong));
    }

    @Test
    void writeLaunchContextBeforeConnect_repairsScenarioPathInIni(@TempDir Path tmp) throws Exception {
        Path iniPath = tmp.resolve("RAP設定.ini");
        Files.writeString(
                iniPath,
                """
                起動プログラム番号=1
                2="Z:\\portable\\Aladdin_RPA_Studio.exe" --scenario \\\\192.168.0.101\\share\\002 加工G\\工程別.ardrpa
                """,
                StandardCharsets.UTF_8);

        String repairedArgs =
                RdpRemoteLauncherIni.mergeEternalFlag(
                        RdpRemoteLauncherIni.normalizeScenarioArguments(
                                "--scenario \\\\192.168.0.101\\share\\002 加工G\\工程別.ardrpa"),
                        false);
        RdpRemoteLauncherIni.writeLaunchContextBeforeConnect(
                iniPath,
                2,
                "Z:\\portable\\Aladdin_RPA_Studio.exe",
                repairedArgs,
                RdpSessionEndAction.SIGN_OUT);

        String text = Files.readString(iniPath, StandardCharsets.UTF_8);
        assertTrue(text.contains("起動プログラム番号=2"));
        assertTrue(text.contains("002  加工G"), "saved ini=" + text);
    }

    @Test
    void load_repairsCollapsedScenarioPathOnRead(@TempDir Path tmp) throws Exception {
        Path iniPath = tmp.resolve("RAP設定.ini");
        Files.writeString(
                iniPath,
                """
                起動プログラム番号=2
                2="Z:\\portable\\Aladdin_RPA_Studio.exe" --scenario \\\\192.168.0.101\\share\\002 加工G\\工程別.ardrpa
                """,
                StandardCharsets.UTF_8);

        RdpRemoteLauncherIni loaded = RdpRemoteLauncherIni.load(iniPath);
        assertTrue(
                loaded.getSlotCommand(2).arguments().contains("002  加工G"),
                "arguments=" + loaded.getSlotCommand(2).arguments());
    }

    @Test
    void loadAndSave_roundTripScenarioArgumentsFromUi(@TempDir Path tmp) throws Exception {
        String uiArgs =
                "--scenario \"\\\\192.168.0.101\\share\\002  加工G\\工程別.ardrpa\"";
        RdpRemoteLauncherIni ini = new RdpRemoteLauncherIni();
        ini.setSelectedSlot(1);
        ini.setSlotCommand(
                1,
                "Z:\\portable\\Aladdin_RPA_Studio.exe",
                RdpRemoteLauncherIni.mergeEternalFlag(
                        RdpRemoteLauncherIni.normalizeScenarioArguments(uiArgs), false));
        Path iniPath = tmp.resolve("RAP設定.ini");
        ini.save(iniPath);

        RdpRemoteLauncherIni loaded = RdpRemoteLauncherIni.load(iniPath);
        assertTrue(
                loaded.getSlotCommand(1).arguments().contains("工程別.ardrpa"),
                "arguments=" + loaded.getSlotCommand(1).arguments());
        assertEquals(
                uiArgs.replace("\"", ""),
                RdpRemoteLauncherIni.argumentsForUiDisplayWithoutManagedFlags(
                        loaded.getSlotCommand(1).arguments()));
    }

    @Test
    void normalizeScenarioArguments_convertsBarePath() {
        assertEquals(
                "--scenario \\\\server\\share\\a.ardrpa",
                RdpRemoteLauncherIni.normalizeScenarioArguments("\\\\server\\share\\a.ardrpa"));
    }

    @Test
    void setSlotCommand_stripsSurroundingQuotesFromProgram(@TempDir Path tmp) throws Exception {
        RdpRemoteLauncherIni ini = new RdpRemoteLauncherIni();
        ini.setSlotCommand(1, "\"Z:\\app\\Aladdin_RPA_Studio.exe\"", "");
        assertEquals("Z:\\app\\Aladdin_RPA_Studio.exe", ini.getSlotCommand(1).executable());
    }

    @Test
    void formatScenarioArgument_usesScenarioFlag() {
        assertEquals(
                "--scenario C:\\tmp\\a.ardrpa",
                RdpRemoteLauncherIni.formatScenarioArgument("C:\\tmp\\a.ardrpa"));
    }

    @Test
    void argumentsForUiDisplayWithoutEternal_stripsFlag() {
        assertEquals(
                "--scenario \\\\server\\share\\a.ardrpa",
                RdpRemoteLauncherIni.argumentsForUiDisplayWithoutEternal(
                        "--scenario \\\\server\\share\\a.ardrpa --eternal"));
        assertTrue(
                RdpRemoteLauncherIni.hasEternalFlag(
                        "--scenario \\\\server\\share\\a.ardrpa --eternal"));
    }

    @Test
    void maxRpaProfileNumber_ignoresSignOutSlot() {
        assertEquals(
                5,
                RdpRemoteLauncherIni.maxRpaProfileNumber(
                        java.util.List.of(
                                RdpRemoteLauncherIni.SLOT_SIGN_OUT, 1, 2, 3, 4, 5)));
        assertEquals(3, RdpRemoteLauncherIni.maxRpaProfileNumber(java.util.List.of(99, 1, 2, 3)));
    }

    @Test
    void visibleSlotCount_reflectsHighestDefinedSlot() {
        RdpRemoteLauncherIni ini = new RdpRemoteLauncherIni();
        assertEquals(RdpRemoteLauncherIni.DEFAULT_INITIAL_RPA_PROFILE_ROWS, ini.visibleSlotCount());
        ini.setSlotCommand(5, "Z:\\app.exe", "");
        assertEquals(5, ini.visibleSlotCount());
    }
}
