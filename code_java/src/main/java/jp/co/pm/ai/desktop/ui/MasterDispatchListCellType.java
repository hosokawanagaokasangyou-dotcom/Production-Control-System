package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.List;

import org.controlsfx.control.spreadsheet.SpreadsheetCellEditor;
import org.controlsfx.control.spreadsheet.SpreadsheetCell;
import org.controlsfx.control.spreadsheet.SpreadsheetCellType;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

/**
 * 配台マスタの LIST セル。ControlsFX 既定のセル内 ComboBox は行高・列幅を押し広げて
 * 固定見出しと本体がずれるため、Popup ピッカーを使う。
 */
public final class MasterDispatchListCellType extends SpreadsheetCellType.ListType {

    public MasterDispatchListCellType(List<String> items) {
        super(items);
    }

    static boolean isPopupListCell(SpreadsheetCell cell) {
        return cell != null && cell.getCellType() instanceof MasterDispatchListCellType;
    }

    @Override
    public SpreadsheetCellEditor createEditor(SpreadsheetView view) {
        List<String> choices = new ArrayList<>();
        if (items != null) {
            for (Object item : items) {
                choices.add(item == null ? "" : item.toString());
            }
        }
        return new MasterDispatchListPopupEditor(view, choices);
    }
}
