package jp.co.pm.ai.desktop.ui;

import org.controlsfx.control.spreadsheet.SpreadsheetCellEditor;
import org.controlsfx.control.spreadsheet.SpreadsheetCellType;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

/** 配台マスタ加工速度など、範囲付き小数のセル型（テンキー編集）。 */
public final class MasterDispatchDecimalCellType extends SpreadsheetCellType.StringType {

    private final double min;
    private final double max;
    private final int fractionDigits;

    public MasterDispatchDecimalCellType(double min, double max, int fractionDigits) {
        this.min = min;
        this.max = max;
        this.fractionDigits = Math.max(0, fractionDigits);
    }

    @Override
    public SpreadsheetCellEditor createEditor(SpreadsheetView view) {
        return new MasterDispatchDecimalKeypadEditor(view, min, max, fractionDigits);
    }

    @Override
    public boolean match(Object value, Object... options) {
        if (value == null) {
            return true;
        }
        String s = value.toString().strip();
        return s.isEmpty()
                || MasterDispatchSheetEditRules.isDecimalInRange(s, min, max, fractionDigits);
    }
}
