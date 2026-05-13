package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** {@link PlanInputTabularIo} と Python {@code _resolve_tabular_sheet_name_calamine} の単一シート解決の整合。 */
class PlanInputTabularIoTest {

    @Test
    void normSheetKey_fullWidthLowLineMatchesAsciiUnderscore() {
        String ascii = "配台計画_タスク入力";
        String full = "配台計画\uFF3Fタスク入力";
        assertEquals(PlanInputTabularIo.normSheetKey(ascii), PlanInputTabularIo.normSheetKey(full));
    }

    @Test
    void readWithResolvedSheet_nfkResolvesToActualSheetName(@TempDir Path root) throws Exception {
        Path xlsx = root.resolve("nfk.xlsx");
        String full = "配台計画\uFF3Fタスク入力";
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("other");
            wb.createSheet(full);
            try (var out = Files.newOutputStream(xlsx)) {
                wb.write(out);
            }
        }
        PlanInputTabularIo.TabularRead r =
                PlanInputTabularIo.readWithResolvedSheet(xlsx, "配台計画_タスク入力");
        assertEquals(full, r.resolvedSheetName());
    }

    @Test
    void readWithResolvedSheet_singleSheetFallback_matchesPythonCalamine(@TempDir Path root) throws Exception {
        Path xlsx = root.resolve("plan_input_tasks.xlsx");
        PlanInputTabularIo.write(
                xlsx,
                "タスク一覧",
                new PlanInputTabularIo.TabularSheet(
                        List.of("依頼NO", "工程名"), List.of(List.of("T1", "加工A"))));

        PlanInputTabularIo.TabularRead r =
                PlanInputTabularIo.readWithResolvedSheet(xlsx, "配台計画_タスク入力");
        assertEquals("タスク一覧", r.resolvedSheetName());
        assertFalse(r.tabular().headers().isEmpty());
        assertEquals("依頼NO", r.tabular().headers().get(0));
        assertEquals("T1", r.tabular().rows().get(0).get(0));
    }

    @Test
    void read_csv_returnsEmptyResolvedSheetName(@TempDir Path root) throws Exception {
        Path csv = root.resolve("p.csv");
        Files.writeString(csv, "a,b\n1,2\n", StandardCharsets.UTF_8);
        PlanInputTabularIo.TabularRead r = PlanInputTabularIo.readWithResolvedSheet(csv, "ignored");
        assertEquals("", r.resolvedSheetName());
        assertEquals("1", r.tabular().rows().get(0).get(0));
    }
}
