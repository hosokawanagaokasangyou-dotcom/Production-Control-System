package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jp.co.pm.ai.desktop.dispatch.DispatchAladdinEntrySheetBuilder;
import jp.co.pm.ai.desktop.reconciliation.JuchuTransferValueNormalizer;

/**
 * アラジン入力用配台計画 Excel から日付セルの（シス計）を読み取る。
 */
public final class AladdinEntryDispatchPlanWorkbookReader {

    private static final String COL_TID = "依頼NO";
    private static final String COL_PROCESS = "工程名";

    private AladdinEntryDispatchPlanWorkbookReader() {}

    /**
     * データ行のシス計だけを返す（日加工合計数行・空セル・シス計 0 は除外）。
     *
     * @param referenceDate 年なし日付見出し（{@code M/d}）の年推定に使う
     */
    public static List<AladdinEntryDispatchPlanIdentityCheck.SystemQty> readSystemQtys(
            Path xlsx, LocalDate referenceDate) throws IOException {
        if (xlsx == null || !Files.isRegularFile(xlsx)) {
            throw new IOException("配台計画 Excel がありません: " + xlsx);
        }
        LocalDate ref = referenceDate != null ? referenceDate : LocalDate.now();
        List<AladdinEntryDispatchPlanIdentityCheck.SystemQty> out = new ArrayList<>();
        try (InputStream in = Files.newInputStream(xlsx);
                XSSFWorkbook wb = new XSSFWorkbook(in)) {
            for (int s = 0; s < wb.getNumberOfSheets(); s++) {
                Sheet sh = wb.getSheetAt(s);
                if (sh == null) {
                    continue;
                }
                String sheetName = sh.getSheetName();
                if (sheetName == null
                        || sheetName.isBlank()
                        || "データなし".equals(sheetName)) {
                    continue;
                }
                Row header = sh.getRow(0);
                if (header == null) {
                    continue;
                }
                int tidCol = findHeaderCol(header, COL_TID);
                int procCol = findHeaderCol(header, COL_PROCESS);
                if (tidCol < 0) {
                    continue;
                }
                Map<Integer, LocalDate> dateCols = dateColumns(header, ref);
                if (dateCols.isEmpty()) {
                    continue;
                }
                int lastRow = sh.getLastRowNum();
                for (int r = 1; r <= lastRow; r++) {
                    Row row = sh.getRow(r);
                    if (row == null) {
                        continue;
                    }
                    String tid = ExcelCellReadSupport.cellToDisplayString(row.getCell(tidCol)).strip();
                    if (tid.isEmpty()
                            || DispatchAladdinEntryWorkbookExporter.DAILY_PROCESSING_TOTAL_LABEL.equals(
                                    tid)) {
                        continue;
                    }
                    String process =
                            procCol >= 0
                                    ? ExcelCellReadSupport.cellToDisplayString(row.getCell(procCol))
                                            .strip()
                                    : "";
                    for (Map.Entry<Integer, LocalDate> e : dateCols.entrySet()) {
                        double qty = parseSystemQty(cellText(row.getCell(e.getKey())));
                        if (Math.abs(qty) <= DispatchAladdinEntrySheetBuilder.QTY_MATCH_EPS) {
                            continue;
                        }
                        out.add(
                                new AladdinEntryDispatchPlanIdentityCheck.SystemQty(
                                        sheetName, tid, process, e.getValue(), qty));
                    }
                }
            }
        }
        return List.copyOf(out);
    }

    static double parseSystemQty(String cellText) {
        if (cellText == null || cellText.isBlank()) {
            return 0d;
        }
        String prefix = DispatchAladdinEntrySheetBuilder.SYSTEM_LINE_PREFIX;
        int idx = cellText.indexOf(prefix);
        if (idx < 0) {
            return 0d;
        }
        String rest = cellText.substring(idx + prefix.length());
        int nl = rest.indexOf('\n');
        if (nl >= 0) {
            rest = rest.substring(0, nl);
        }
        return JuchuTransferValueNormalizer.normalizeNumeric(rest);
    }

    private static Map<Integer, LocalDate> dateColumns(Row header, LocalDate ref) {
        Map<Integer, LocalDate> out = new LinkedHashMap<>();
        short last = header.getLastCellNum();
        for (int c = 0; c < last; c++) {
            String text = ExcelCellReadSupport.cellToDisplayString(header.getCell(c));
            LocalDate d = JuchuTransferValueNormalizer.parseLocalDate(text, ref);
            if (d != null) {
                out.put(c, d);
            }
        }
        return out;
    }

    private static int findHeaderCol(Row header, String title) {
        short last = header.getLastCellNum();
        for (int c = 0; c < last; c++) {
            String text = ExcelCellReadSupport.cellToDisplayString(header.getCell(c)).strip();
            if (title.equals(text)) {
                return c;
            }
        }
        return -1;
    }

    private static String cellText(Cell cell) {
        return ExcelCellReadSupport.cellToDisplayString(cell);
    }
}
