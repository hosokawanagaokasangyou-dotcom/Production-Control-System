package jp.co.pm.ai.planning.stage2.core;

import java.io.IOException;
import java.nio.file.Path;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * マスタブックに need／機械カレンダー等のシートが存在するかの軽量検出（Python load 前段の足場）。
 */
public record Stage2MasterSheetProbe(
        boolean hasNeedSheet,
        boolean hasMachineCalendarSheet,
        int sheetCount) {

    public static Stage2MasterSheetProbe scan(Path master) throws IOException {
        try (Workbook wb = WorkbookFactory.create(master.toFile())) {
            boolean need = false;
            boolean cal = false;
            int n = wb.getNumberOfSheets();
            for (int i = 0; i < n; i++) {
                String name = wb.getSheetName(i);
                if (name != null) {
                    if (name.contains("need")) {
                        need = true;
                    }
                    if (name.contains("機械カレンダー")) {
                        cal = true;
                    }
                }
            }
            return new Stage2MasterSheetProbe(need, cal, n);
        }
    }
}
