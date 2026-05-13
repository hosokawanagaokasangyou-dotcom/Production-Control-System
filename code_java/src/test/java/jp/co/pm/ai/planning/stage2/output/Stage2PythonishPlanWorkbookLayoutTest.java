package jp.co.pm.ai.planning.stage2.output;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.io.PlanInputTabularIo;

class Stage2PythonishPlanWorkbookLayoutTest {

    @Test
    void sheetNamesAndOrderMatchPythonStage2(@TempDir Path tmp) throws Exception {
        Path x = tmp.resolve("計画test.xlsx");
        PlanInputTabularIo.TabularSheet tab =
                new PlanInputTabularIo.TabularSheet(List.of("依頼NO"), List.of(List.of("X1")));
        Stage2PythonishPlanWorkbookLayout.write(x, tab, List.of("工程1+機1"), List.of("山田"));
        try (InputStream in = Files.newInputStream(x);
                Workbook wb = WorkbookFactory.create(in)) {
            assertEquals(Stage2PythonishPlanWorkbookLayout.PLAN_SHEET_ORDER.size(), wb.getNumberOfSheets());
            for (int i = 0; i < Stage2PythonishPlanWorkbookLayout.PLAN_SHEET_ORDER.size(); i++) {
                assertEquals(
                        Stage2PythonishPlanWorkbookLayout.PLAN_SHEET_ORDER.get(i),
                        wb.getSheetAt(i).getSheetName());
            }
            var taskSheet = wb.getSheet("結果_タスク一覧");
            assertEquals(
                    Stage2ResultTaskListCanonicalHeaders.DEFAULT_ORDER_NO_HISTORY.size(),
                    (int) taskSheet.getRow(0).getLastCellNum());
        }
    }
}
