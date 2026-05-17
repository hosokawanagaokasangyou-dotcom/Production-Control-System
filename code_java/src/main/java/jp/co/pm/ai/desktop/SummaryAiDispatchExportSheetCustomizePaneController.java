package jp.co.pm.ai.desktop;

import java.util.ArrayList;
import java.util.List;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.ListView;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;

import jp.co.pm.ai.desktop.config.SummaryAiDispatchExportPrefs;
import jp.co.pm.ai.desktop.io.SummaryAiDispatchExportColumnSupport;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;

/** サマリ Excel 出力: 1 シート分の見出し列数・非日付列順。 */
public final class SummaryAiDispatchExportSheetCustomizePaneController {

    @FXML
    private Spinner<Integer> frozenColumnSpinner;

    @FXML
    private ListView<String> columnOrderList;

    private SummaryAiDispatchExportPrefs.SheetKey sheetKey;
    private MainShellController shell;

    void bind(MainShellController shell, SummaryAiDispatchExportPrefs.SheetKey key) {
        this.shell = shell;
        this.sheetKey = key;
        if (frozenColumnSpinner != null) {
            frozenColumnSpinner.setValueFactory(
                    new SpinnerValueFactory.IntegerSpinnerValueFactory(0, 200, key.defaultFrozenColumns()));
        }
        reloadFromStore();
    }

    void reloadFromStore() {
        if (sheetKey == null) {
            return;
        }
        SummaryAiDispatchExportPrefs.ExportPrefs prefs = SummaryAiDispatchExportPrefs.load();
        SummaryAiDispatchExportPrefs.SheetPrefs sp = prefs.sheet(sheetKey);
        if (frozenColumnSpinner != null) {
            frozenColumnSpinner.getValueFactory().setValue(sp.frozenColumnCount());
        }
        if (columnOrderList != null) {
            columnOrderList.setItems(FXCollections.observableArrayList(sp.nonDateColumnOrder()));
        }
    }

    @FXML
    private void onMoveUpAction() {
        moveSelected(-1);
    }

    @FXML
    private void onMoveDownAction() {
        moveSelected(1);
    }

    @FXML
    private void onClearColumnOrderAction() {
        if (columnOrderList != null) {
            columnOrderList.getItems().clear();
        }
    }

    @FXML
    private void onImportFromUiAction() {
        if (sheetKey == null) {
            return;
        }
        TableColumnOrderPersistence.TableId tid = tableIdForSheet(sheetKey);
        if (tid == null) {
            if (shell != null) {
                shell.appendLog("[summary-export] このシートに対応する UI 表がありません: " + sheetKey.sheetName());
            }
            return;
        }
        List<TableColumnOrderPersistence.ColumnSpec> lay = TableColumnOrderPersistence.loadLayout(tid);
        List<String> titles = new ArrayList<>();
        for (TableColumnOrderPersistence.ColumnSpec spec : lay) {
            if (spec != null && spec.title() != null && !spec.title().isBlank()) {
                if (!SummaryAiDispatchExportColumnSupport.isDateColumnHeader(sheetKey, spec.title())) {
                    titles.add(spec.title());
                }
            }
        }
        if (columnOrderList != null) {
            columnOrderList.setItems(FXCollections.observableArrayList(titles));
        }
        if (shell != null) {
            shell.appendLog(
                    "[summary-export] "
                            + sheetKey.sheetName()
                            + " ← UI 列順 "
                            + titles.size()
                            + " 列（日付列除く）");
        }
    }

    @FXML
    private void onSaveSheetAction() {
        if (sheetKey == null) {
            return;
        }
        int frozen =
                frozenColumnSpinner != null && frozenColumnSpinner.getValue() != null
                        ? frozenColumnSpinner.getValue()
                        : sheetKey.defaultFrozenColumns();
        List<String> order =
                columnOrderList != null
                        ? new ArrayList<>(columnOrderList.getItems())
                        : List.of();
        SummaryAiDispatchExportPrefs.saveSheetPrefs(
                sheetKey, new SummaryAiDispatchExportPrefs.SheetPrefs(frozen, order));
        if (shell != null) {
            shell.appendLog(
                    "[summary-export] 保存: "
                            + sheetKey.sheetName()
                            + " 見出し列="
                            + frozen
                            + " 非日付列="
                            + order.size());
        }
    }

    private void moveSelected(int delta) {
        if (columnOrderList == null) {
            return;
        }
        int idx = columnOrderList.getSelectionModel().getSelectedIndex();
        if (idx < 0) {
            return;
        }
        int next = idx + delta;
        ObservableList<String> items = columnOrderList.getItems();
        if (next < 0 || next >= items.size()) {
            return;
        }
        String item = items.remove(idx);
        items.add(next, item);
        columnOrderList.getSelectionModel().select(next);
    }

    private static TableColumnOrderPersistence.TableId tableIdForSheet(
            SummaryAiDispatchExportPrefs.SheetKey key) {
        return switch (key) {
            case MAIN_COMPARE -> TableColumnOrderPersistence.TableId.DELIVERY_CALENDAR_MAIN;
            case DISPATCH -> TableColumnOrderPersistence.TableId.RESULT_DISPATCH_TABLE;
            case ACTUALS -> TableColumnOrderPersistence.TableId.PROCESSING_ACTUALS_DETAIL_RAW;
            case ALADDIN -> TableColumnOrderPersistence.TableId.ALADDIN_PROCESSING_PLAN_RAW;
        };
    }
}
