package jp.co.pm.ai.desktop.io;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.model.CalculationChain;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCalcPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbook;

import java.io.IOException;
import java.io.OutputStream;

/**
 * xlsx/xlsm を POI で保存する際の Open XML 整合処理。
 *
 * <p>数式セルの削除・値への置換（{@code removeCell} + {@code createCell} 等）後に
 * {@code /xl/calcChain.xml} が実シートと乖離すると、Excel が「ファイルの一部が壊れている」
 * と報告し当該パーツを削除して修復することがある。計算チェーンは最適化用のみなので、
 * 保存前に除去し、次回 Excel 起動時にフル再計算させる（Microsoft / Open XML SDK 推奨）。
 */
public final class PoiWorkbookSaver {

    private PoiWorkbookSaver() {}

    /**
     * {@link Workbook#write(OutputStream)} の直前に呼ぶ。xlsm/xlsx の {@link XSSFWorkbook} のみ処理。
     */
    public static void prepareBeforeWrite(Workbook workbook) {
        if (!(workbook instanceof XSSFWorkbook xssf)) {
            return;
        }
        removeCalculationChain(xssf);
        requestFullRecalculationOnLoad(xssf);
    }

    public static void write(Workbook workbook, OutputStream out) throws IOException {
        prepareBeforeWrite(workbook);
        workbook.write(out);
    }

    private static void removeCalculationChain(XSSFWorkbook workbook) {
        CalculationChain chain = workbook.getCalculationChain();
        if (chain == null) {
            return;
        }
        try {
            workbook.getPackage().removePart(chain.getPackagePart());
        } catch (Exception ex) {
            // パーツが既に無い・参照のみ残る場合は無視（保存は続行）
        }
    }

    private static void requestFullRecalculationOnLoad(XSSFWorkbook workbook) {
        CTWorkbook ct = workbook.getCTWorkbook();
        if (ct == null) {
            return;
        }
        CTCalcPr calcPr = ct.isSetCalcPr() ? ct.getCalcPr() : ct.addNewCalcPr();
        calcPr.setFullCalcOnLoad(true);
    }
}
