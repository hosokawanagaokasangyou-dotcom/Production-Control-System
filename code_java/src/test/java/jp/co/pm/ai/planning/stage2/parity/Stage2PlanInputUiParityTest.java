package jp.co.pm.ai.planning.stage2.parity;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.io.PlanInputTabularIo;

class Stage2PlanInputUiParityTest {

    @Test
    void compare_columnOrderPermutation_ok(@TempDir Path root) throws Exception {
        Path x = root.resolve("in.xlsx");
        PlanInputTabularIo.write(
                x,
                "S",
                new PlanInputTabularIo.TabularSheet(
                        List.of("A", "B"), List.of(List.of("1", "2"), List.of("3", "4"))));
        PlanInputTabularIo.TabularSheet ui =
                new PlanInputTabularIo.TabularSheet(
                        List.of("B", "A"), List.of(List.of("2", "1"), List.of("4", "3")));
        Stage2ParityCheckResult r = Stage2PlanInputUiParity.compareUiToDisk(ui, x, "S");
        assertTrue(r.identical());
    }

    @Test
    void compare_cellMismatch(@TempDir Path root) throws Exception {
        Path x = root.resolve("in.xlsx");
        PlanInputTabularIo.write(
                x,
                "S",
                new PlanInputTabularIo.TabularSheet(
                        List.of("A"), List.of(List.of("x"))));
        PlanInputTabularIo.TabularSheet ui =
                new PlanInputTabularIo.TabularSheet(List.of("A"), List.of(List.of("y")));
        Stage2ParityCheckResult r = Stage2PlanInputUiParity.compareUiToDisk(ui, x, "S");
        assertFalse(r.identical());
    }
}
