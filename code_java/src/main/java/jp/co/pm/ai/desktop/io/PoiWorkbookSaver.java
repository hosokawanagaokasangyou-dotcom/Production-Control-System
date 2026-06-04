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
    public static void write(Workbook workbook, OutputStream out) throws IOException {
        prepareBeforeWrite(workbook, true);
        workbook.write(out);
    }

    /**
     * calcChain パーツ削除を行わず保存する（{@code partName} 等で標準保存が失敗する修復済みブック向け）。
     */
    public static void writeLenient(Workbook workbook, OutputStream out) throws IOException {
        prepareBeforeWrite(workbook, false);
        workbook.write(out);
    }

    static boolean isPartNameFailure(Throwable ex) {
        for (Throwable t = ex; t != null; t = t.getCause()) {
            if (t instanceof IllegalArgumentException ia && "partName".equals(ia.getMessage())) {
                return true;
            }
        }
        return false;
    }

    private static void prepareBeforeWrite(Workbook workbook, boolean removeCalcChain) {
        if (!(workbook instanceof XSSFWorkbook xssf)) {
            return;
        }
        if (removeCalcChain) {
            removeCalculationChain(xssf);
        }
        requestFullRecalculationOnLoad(xssf);
        xssf.setForceFormulaRecalculation(true);
    }

    /** @deprecated 互換のため残す。{@link #write(Workbook, OutputStream)} と同等。 */
    @Deprecated
    public static void prepareBeforeWrite(Workbook workbook) {
        prepareBeforeWrite(workbook, true);
    }

    private static void removeCalculationChain(XSSFWorkbook workbook) {
        CalculationChain chain = workbook.getCalculationChain();
        if (chain == null) {
            return;
        }
        try {
            detachCalculationChainRelation(workbook, chain);
        } catch (Exception ignored) {
            // calcChain 参照のみ残るブックは無視（保存は続行）
        }
    }

    /**
     * {@link XSSFWorkbook} 内部の calcChain 参照を外す（{@code onSheetDelete} と同様に
     * {@code removeRelation} のみ。{@code removePart} 直叩きは {@code write()} 時の二重削除で
     * {@code partName} になる）。
     */
    private static void detachCalculationChainRelation(XSSFWorkbook workbook, CalculationChain chain)
            throws ReflectiveOperationException {
        java.lang.reflect.Method removeRelation =
                org.apache.poi.ooxml.POIXMLDocumentPart.class.getDeclaredMethod(
                        "removeRelation", org.apache.poi.ooxml.POIXMLDocumentPart.class);
        removeRelation.setAccessible(true);
        removeRelation.invoke(workbook, chain);
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
