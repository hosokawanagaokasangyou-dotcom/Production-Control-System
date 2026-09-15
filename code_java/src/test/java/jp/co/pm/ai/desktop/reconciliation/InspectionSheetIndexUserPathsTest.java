package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;
import java.time.LocalDate;
import java.util.List;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class InspectionSheetIndexUserPathsTest {

    @Test
    void relocateBoxHome_rewritesOtherUsersBoxPrefix() {
        Path home = Path.of("C:\\Users\\bob");
        String published =
                "C:\\Users\\0585\\Box\\長岡産業\\後加工検査表\\国分工場\\2026_C8-9.xlsx";
        String got = InspectionSheetIndexUserPaths.relocateBoxHome(published, home);
        assertEquals(
                Path.of("C:\\Users\\bob\\Box\\長岡産業\\後加工検査表\\国分工場\\2026_C8-9.xlsx")
                        .normalize()
                        .toString(),
                Path.of(got).normalize().toString());
    }

    @Test
    void relocateBoxHome_leavesUncAndNonBoxPaths() {
        Path home = Path.of("C:\\Users\\bob");
        String unc =
                "\\\\192.168.0.101\\共有フォルダ\\湖南工場\\湖南共有\\生産管理システム\\a.xlsx";
        assertEquals(unc, InspectionSheetIndexUserPaths.relocateBoxHome(unc, home));
        String localTmp = "C:\\tmp\\inspection\\a.xlsx";
        assertEquals(localTmp, InspectionSheetIndexUserPaths.relocateBoxHome(localTmp, home));
    }

    @Test
    void load_relocatesBoxPathToCurrentUserHome(@TempDir Path tmp) throws Exception {
        String prior = System.getProperty("user.home");
        Path home = tmp.resolve("Users").resolve("bob");
        System.setProperty("user.home", home.toString());
        try {
            Path csv = tmp.resolve("KOKUBU.csv");
            InspectionSheetIndexStore.Row row =
                    new InspectionSheetIndexStore.Row(
                            "C8-9",
                            LocalDate.of(2026, 8, 18),
                            "2026-08",
                            "C:\\Users\\0585\\Box\\長岡産業\\後加工検査表\\国分工場\\a.xlsx",
                            "a.xlsx",
                            1L,
                            2L,
                            "t");
            InspectionSheetIndexStore.save(csv, List.of(row));
            List<InspectionSheetIndexStore.Row> loaded = InspectionSheetIndexStore.load(csv);
            assertEquals(1, loaded.size());
            Path got = Path.of(loaded.get(0).filePath()).normalize();
            Path expected =
                    home.resolve("Box")
                            .resolve("長岡産業")
                            .resolve("後加工検査表")
                            .resolve("国分工場")
                            .resolve("a.xlsx")
                            .normalize();
            assertEquals(expected, got);
            assertTrue(got.toString().contains("bob"));
        } finally {
            System.setProperty("user.home", prior);
        }
    }
}
