package jp.co.pm.ai.kouchin.verify;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class CheckMatomeTorayDiffTest {

    @TempDir
    Path tempDir;

    @Test
    @DisplayName("同一契約のまとめAAを合算し、①に無い契約は出さない")
    void aggregatesSameKeiyakuAndSkipsMissingToray() {
        List<MatomeCheckResult.MatomeRow> rows = CheckMatome.diffTorayMatome(List.of(
                new CheckMatome.TorayEntry("東レまとめ!166 ↔ 東レV.C!24", "V9-9", "192265M", 480930, 480930),
                new CheckMatome.TorayEntry("東レまとめ!167 ↔ 東レV.C!25", "V9-9", "192-265M", 100, 100),
                new CheckMatome.TorayEntry("東レまとめ!10 ↔ 東レT!8", "T9-1", "199999A", 50, 50)),
                Map.of("192265M", 307808.0),
                0.5);

        assertEquals(1, rows.size());
        MatomeCheckResult.MatomeRow row = rows.get(0);
        assertEquals(Judge.TORAY_DIFF, row.judge());
        assertEquals(307808.0, row.torayAmount(), 0.001);
        assertEquals(481030.0, row.matomeAa(), 0.001);
        assertEquals(307808.0 - 481030.0, row.diff(), 0.001);
        assertTrue(row.place().contains("東レまとめ!166"), row.place());
        assertTrue(row.place().contains("東レV.C!24"), row.place());
        assertTrue(row.place().contains(" / "), row.place());
        assertEquals("V9-9", row.irai());
    }

    @Test
    @DisplayName("まとめGの元シート参照は式異常にせず、①差額は要修正に含めない")
    void sourceRefFormulaAndTorayDiffAreNotInternalErrors() throws Exception {
        Path file = tempDir.resolve("nagaoka.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            for (String name : new String[] {"東レまとめ", "東レT", "東レV.C", "東レY", "東レW.E"}) {
                wb.createSheet(name);
            }
            fillSource(wb.getSheet("東レT"), 6, "V9", "9", "192265M", 10, 3);
            fillSource(wb.getSheet("東レT"), 7, "V9", "9", "192265M", 4, 5);
            fillSource(wb.getSheet("東レT"), 8, "T9", "1", "192266H", 2, 2);
            fillMatome(wb.getSheet("東レまとめ"), 6, "東レT", 6);
            fillMatome(wb.getSheet("東レまとめ"), 7, "東レT", 7);
            fillMatome(wb.getSheet("東レまとめ"), 8, "東レT", 8);
            XSSFFormulaEvaluator.evaluateAllFormulaCells(wb);
            try (var out = Files.newOutputStream(file)) {
                wb.write(out);
            }
        }

        MatomeCheckResult result = CheckMatome.check(file, 0.5, Map.of(
                "192265M", 307808.0,
                "192266H", 4.0));

        assertTrue(result.rows().stream().noneMatch(r ->
                Judge.BAD_FORMULA.equals(r.judge()) && r.place().startsWith("東レまとめ!G")),
                result.rows().toString());
        assertEquals(0, result.errorCount(), result.rows().toString());
        assertEquals(1, result.torayDiffCount(), result.rows().toString());
        MatomeCheckResult.MatomeRow diff = result.rows().stream()
                .filter(r -> Judge.TORAY_DIFF.equals(r.judge()))
                .findFirst()
                .orElseThrow();
        assertEquals("192265M", Norm.keiyaku(diff.keiyaku()));
        assertEquals(307808.0, diff.torayAmount(), 0.001);
        assertEquals(50.0, diff.matomeAa(), 0.001);
        assertTrue(diff.place().contains("東レまとめ!6"), diff.place());
        assertTrue(diff.place().contains("東レT!7"), diff.place());
        assertEquals(0, result.noticeCount(), result.rows().toString());
    }

    private static void fillSource(Sheet sheet, int row1, String a, String b, String keiyaku, double e, double h) {
        Row row = sheet.createRow(row1 - 1);
        row.createCell(0).setCellValue(a);
        row.createCell(1).setCellValue(b);
        row.createCell(2).setCellValue(keiyaku);
        row.createCell(4).setCellValue(e);
        row.createCell(7).setCellValue(h);
        row.createCell(6).setCellFormula("SUM(H" + row1 + ":X" + row1 + ")");
        row.createCell(26).setCellFormula("ROUNDUP($E" + row1 + "*G" + row1 + ",0)");
    }

    private static void fillMatome(Sheet sheet, int row1, String srcSheet, int srcRow) {
        Row row = sheet.createRow(row1 - 1);
        for (int c = 0; c <= 23; c++) {
            String letter = CellReference.convertNumToColString(c);
            if ("G".equals(letter)) {
                continue;
            }
            Cell cell = row.createCell(c);
            cell.setCellFormula("'" + srcSheet + "'!" + letter + srcRow);
        }
        row.createCell(6).setCellFormula("'" + srcSheet + "'!G" + srcRow);
        row.createCell(25).setCellFormula("ROUNDUP($D" + row1 + "*G" + row1 + ",0)");
        row.createCell(26).setCellFormula("ROUNDUP($E" + row1 + "*G" + row1 + ",0)");
    }
}
