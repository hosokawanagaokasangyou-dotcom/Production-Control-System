package jp.co.pm.ai.desktop;

import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.TextField;

import jp.co.pm.ai.desktop.config.DesktopTheme;
import jp.co.pm.ai.desktop.config.SummaryAiDispatchExportPrefs;

/** メインシェル「サマリ_AI配台エクセル生成カスタイマイズ」タブ。 */
public final class SummaryAiDispatchExportCustomizeTabController {

    @FXML
    private TextField fontFamilyField;

    @FXML
    private Spinner<Integer> fontSizeSpinner;

    @FXML
    private ComboBox<DesktopTheme> themeCombo;

    @FXML
    private SummaryAiDispatchExportSheetCustomizePaneController mainComparePaneController;

    @FXML
    private SummaryAiDispatchExportSheetCustomizePaneController dispatchPaneController;

    @FXML
    private SummaryAiDispatchExportSheetCustomizePaneController actualsPaneController;

    @FXML
    private SummaryAiDispatchExportSheetCustomizePaneController aladdinPaneController;

    private MainShellController shell;

    void bindShell(MainShellController shell) {
        this.shell = shell;
        if (fontSizeSpinner != null) {
            fontSizeSpinner.setValueFactory(
                    new SpinnerValueFactory.IntegerSpinnerValueFactory(
                            6, 72, SummaryAiDispatchExportPrefs.DEFAULT_FONT_SIZE_PT));
        }
        if (themeCombo != null) {
            themeCombo.getItems().setAll(DesktopTheme.values());
        }
        if (mainComparePaneController != null) {
            mainComparePaneController.bind(shell, SummaryAiDispatchExportPrefs.SheetKey.MAIN_COMPARE);
        }
        if (dispatchPaneController != null) {
            dispatchPaneController.bind(shell, SummaryAiDispatchExportPrefs.SheetKey.DISPATCH);
        }
        if (actualsPaneController != null) {
            actualsPaneController.bind(shell, SummaryAiDispatchExportPrefs.SheetKey.ACTUALS);
        }
        if (aladdinPaneController != null) {
            aladdinPaneController.bind(shell, SummaryAiDispatchExportPrefs.SheetKey.ALADDIN);
        }
        reloadGlobalFromStore();
    }

    private void reloadGlobalFromStore() {
        SummaryAiDispatchExportPrefs.ExportPrefs prefs = SummaryAiDispatchExportPrefs.load();
        if (fontFamilyField != null) {
            fontFamilyField.setText(prefs.fontFamily());
        }
        if (fontSizeSpinner != null) {
            fontSizeSpinner.getValueFactory().setValue(prefs.fontSizePt());
        }
        if (themeCombo != null) {
            themeCombo.setValue(prefs.theme());
        }
    }

    @FXML
    private void onSaveGlobalAction() {
        SummaryAiDispatchExportPrefs.ExportPrefs cur = SummaryAiDispatchExportPrefs.load();
        String font =
                fontFamilyField != null && fontFamilyField.getText() != null
                        ? fontFamilyField.getText().trim()
                        : SummaryAiDispatchExportPrefs.DEFAULT_FONT_FAMILY;
        if (font.isEmpty()) {
            font = SummaryAiDispatchExportPrefs.DEFAULT_FONT_FAMILY;
        }
        int size =
                fontSizeSpinner != null && fontSizeSpinner.getValue() != null
                        ? fontSizeSpinner.getValue()
                        : SummaryAiDispatchExportPrefs.DEFAULT_FONT_SIZE_PT;
        DesktopTheme theme =
                themeCombo != null && themeCombo.getValue() != null
                        ? themeCombo.getValue()
                        : SummaryAiDispatchExportPrefs.DEFAULT_THEME;
        SummaryAiDispatchExportPrefs.save(
                new SummaryAiDispatchExportPrefs.ExportPrefs(font, size, theme, cur.sheets()));
        if (shell != null) {
            shell.appendLog("[summary-export] 共通設定を保存しました（フォント・テーマ）");
        }
    }

    @FXML
    private void onResetAllAction() {
        SummaryAiDispatchExportPrefs.resetToDefaults();
        reloadGlobalFromStore();
        if (mainComparePaneController != null) {
            mainComparePaneController.reloadFromStore();
        }
        if (dispatchPaneController != null) {
            dispatchPaneController.reloadFromStore();
        }
        if (actualsPaneController != null) {
            actualsPaneController.reloadFromStore();
        }
        if (aladdinPaneController != null) {
            aladdinPaneController.reloadFromStore();
        }
        if (shell != null) {
            shell.appendLog("[summary-export] すべて既定値に戻しました");
            Alert ok = new Alert(AlertType.INFORMATION);
            if (shell.primaryStageForDialogs() != null) {
                ok.initOwner(shell.primaryStageForDialogs());
            }
            ok.setTitle("完了");
            ok.setHeaderText(null);
            ok.setContentText("サマリ Excel 出力カスタマイズを既定に戻しました。");
            ok.showAndWait();
        }
    }
}
