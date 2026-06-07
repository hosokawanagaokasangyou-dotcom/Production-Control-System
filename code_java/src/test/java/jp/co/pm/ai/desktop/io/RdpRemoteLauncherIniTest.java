package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

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
    void parseCommandLine_uncWithArgs() {
        String line =
                "\\\\192.168.0.101\\share\\Aladdin_RPA_Studio.exe C:\\Users\\test\\file.ardrpa";
        RdpRemoteLauncherIni.Command cmd = RdpRemoteLauncherIni.parseCommandLine(line);
        assertEquals("\\\\192.168.0.101\\share\\Aladdin_RPA_Studio.exe", cmd.executable());
        assertEquals("C:\\Users\\test\\file.ardrpa", cmd.arguments());
    }

    @Test
    void parseCommandLine_quotedExecutable() {
        RdpRemoteLauncherIni.Command cmd =
                RdpRemoteLauncherIni.parseCommandLine("\"C:\\Program Files\\App\\app.exe\" --flag");
        assertEquals("C:\\Program Files\\App\\app.exe", cmd.executable());
        assertEquals("--flag", cmd.arguments());
    }

    @Test
    void loadAndSave_roundTrip(@TempDir Path tmp) throws Exception {
        Path iniPath = tmp.resolve("RAP設定.ini");
        Files.writeString(
                iniPath,
                """
                起動プログラム番号=2
                1=C:\\Windows\\System32\\notepad.exe
                2=\\\\server\\share\\app.exe arg1
                """,
                StandardCharsets.UTF_8);

        RdpRemoteLauncherIni loaded = RdpRemoteLauncherIni.load(iniPath);
        assertEquals(2, loaded.selectedSlot());
        assertEquals("C:\\Windows\\System32\\notepad.exe", loaded.getSlot(1));

        Path out = tmp.resolve("out.ini");
        loaded.save(out);
        RdpRemoteLauncherIni again = RdpRemoteLauncherIni.load(out);
        assertEquals(2, again.selectedSlot());
        assertEquals("\\\\server\\share\\app.exe arg1", again.getSlot(2));
    }

    @Test
    void validateMessageForSave_emptySelectedSlot() {
        RdpRemoteLauncherIni ini = new RdpRemoteLauncherIni();
        ini.setSelectedSlot(2);
        assertNotNull(ini.validateMessageForSave());
    }

    @Test
    void validateMessageForSave_ok() {
        RdpRemoteLauncherIni ini = new RdpRemoteLauncherIni();
        ini.setSlot(1, "C:\\Windows\\System32\\notepad.exe");
        assertNull(ini.validateMessageForSave());
    }

    @Test
    void parseCommandLine_blankThrows() {
        assertThrows(IllegalArgumentException.class, () -> RdpRemoteLauncherIni.parseCommandLine("  "));
    }
}
