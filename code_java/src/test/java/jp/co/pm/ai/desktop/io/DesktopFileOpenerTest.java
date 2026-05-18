package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.nio.file.Path;

import org.junit.jupiter.api.Test;

class DesktopFileOpenerTest {

    @Test
    void wslPathToWindows_convertsMntDrive() {
        assertEquals(
                "C:\\工程管理AIプロジェクト_JAVA\\code\\サマリ_AI配台.xlsx",
                DesktopFileOpener.wslPathToWindows(
                        Path.of("/mnt/c/工程管理AIプロジェクト_JAVA/code/サマリ_AI配台.xlsx")));
    }

    @Test
    void wslPathToWindows_keepsWindowsAbsolutePaths() {
        assertEquals(
                "C:\\data\\book.xlsx",
                DesktopFileOpener.wslPathToWindows(Path.of("C:/data/book.xlsx")));
    }
}
