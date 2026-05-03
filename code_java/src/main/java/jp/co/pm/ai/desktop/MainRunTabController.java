package jp.co.pm.ai.desktop;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;

import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.text.Font;
import javafx.util.StringConverter;

/** Run/log tab; layout in {@code MainRunTab.fxml}. */
public final class MainRunTabController {

    private static final String DEFAULT_FONT_FAMILY_LABEL = "\u30b7\u30b9\u30c6\u30e0\u65e2\u5b9a";

    private static final List<Double> PRESET_FONT_SIZES =
            List.of(9d, 10d, 11d, 12d, 13d, 14d, 15d, 16d, 18d, 20d, 22d, 24d);

    private MainShellController shell;

    @FXML
    private TextField workbookField;

    @FXML
    private TextField pythonExeField;

    @FXML
    private TextField scriptDirField;

    @FXML
    private TextArea logArea;

    @FXML
    private ComboBox<String> logFontFamilyCombo;

    @FXML
    private ComboBox<Double> logFontSizeCombo;

    @FXML
    private Label statusLabel;

    @FXML
    private Button browseWbButton;

    @FXML
    private Button detectWbButton;

    @FXML
    private Button refreshDirButton;

    @FXML
    private Button stage1RunButton;

    @FXML
    private Button stage2RunButton;

    @FXML
    private Button peekSheetsButton;

    private final AtomicBoolean suppressLogFontEvents = new AtomicBoolean(false);

    @FXML
    private void initialize() {
        List<String> families = new ArrayList<>();
        families.add(DEFAULT_FONT_FAMILY_LABEL);
        List<String> installed = new ArrayList<>(Font.getFamilies());
        Collections.sort(installed);
        families.addAll(installed);
        logFontFamilyCombo.getItems().setAll(families);
        logFontFamilyCombo.getSelectionModel().selectFirst();

        logFontSizeCombo.getItems().setAll(PRESET_FONT_SIZES);
        logFontSizeCombo.setConverter(
                new StringConverter<>() {
                    @Override
                    public String toString(Double object) {
                        if (object == null) {
                            return "";
                        }
                        if (object == Math.rint(object)) {
                            return String.valueOf(object.intValue());
                        }
                        return object.toString();
                    }

                    @Override
                    public Double fromString(String string) {
                        if (string == null || string.isBlank()) {
                            return null;
                        }
                        return Double.valueOf(string.trim());
                    }
                });
        logFontSizeCombo.setValue(14d);

        Runnable onFontUiChange =
                () -> {
                    if (!suppressLogFontEvents.get()) {
                        applyLogAreaFont();
                        if (shell != null) {
                            shell.scheduleDesktopSessionSave();
                        }
                    }
                };
        logFontFamilyCombo.valueProperty().addListener((o, a, b) -> onFontUiChange.run());
        logFontSizeCombo.valueProperty().addListener((o, a, b) -> onFontUiChange.run());

        applyLogAreaFont();
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
    }

    /**
     * Restores font controls from session; must run after FXML {@link #initialize}.
     */
    void applyLogFontFromSession(String family, double sizePoints) {
        suppressLogFontEvents.set(true);
        try {
            if (family != null && !family.isBlank()) {
                if (!logFontFamilyCombo.getItems().contains(family)) {
                    int insertAt = 1;
                    logFontFamilyCombo.getItems().add(insertAt, family);
                }
                logFontFamilyCombo.setValue(family);
            } else {
                logFontFamilyCombo.getSelectionModel().selectFirst();
            }
            double effectiveSize =
                    sizePoints > 0 && Double.isFinite(sizePoints)
                            ? sizePoints
                            : Font.getDefault().getSize();
            if (!logFontSizeCombo.getItems().contains(effectiveSize)) {
                List<Double> extended = new ArrayList<>(logFontSizeCombo.getItems());
                extended.add(effectiveSize);
                Collections.sort(extended);
                logFontSizeCombo.getItems().setAll(extended);
            }
            logFontSizeCombo.setValue(effectiveSize);
        } finally {
            suppressLogFontEvents.set(false);
        }
        applyLogAreaFont();
    }

    String snapshotLogFontFamily() {
        String v = logFontFamilyCombo != null ? logFontFamilyCombo.getValue() : null;
        if (v == null || v.equals(DEFAULT_FONT_FAMILY_LABEL)) {
            return "";
        }
        return v;
    }

    double snapshotLogFontSize() {
        Double v = logFontSizeCombo != null ? logFontSizeCombo.getValue() : null;
        if (v == null || !Double.isFinite(v) || v <= 0) {
            return 0d;
        }
        return v;
    }

    private void applyLogAreaFont() {
        if (logArea == null || logFontFamilyCombo == null || logFontSizeCombo == null) {
            return;
        }
        String choice = logFontFamilyCombo.getValue();
        Double szObj = logFontSizeCombo.getValue();
        double size =
                szObj != null && szObj > 0 && Double.isFinite(szObj)
                        ? szObj
                        : Font.getDefault().getSize();
        if (choice == null || choice.equals(DEFAULT_FONT_FAMILY_LABEL)) {
            logArea.setFont(Font.font(size));
        } else {
            logArea.setFont(Font.font(choice, size));
        }
    }

    @FXML
    private void onBrowseWorkbookButtonAction() {
        shell.pickWorkbook();
    }

    @FXML
    private void onDetectWorkbookButtonAction() {
        workbookField.setText(shell.resolveTaskInputWorkbookFromEnv());
    }

    @FXML
    private void onRefreshScriptDirButtonAction() {
        scriptDirField.setText(shell.resolvePythonScriptDirFromEnv());
    }

    @FXML
    private void onStage1RunButtonAction() {
        shell.triggerStage1();
    }

    @FXML
    private void onStage2RunButtonAction() {
        shell.triggerStage2();
    }

    @FXML
    private void onPeekSheetsButtonAction() {
        shell.triggerPeekSheets();
    }

    TextField getWorkbookField() {
        return workbookField;
    }

    TextField getPythonExeField() {
        return pythonExeField;
    }

    TextField getScriptDirField() {
        return scriptDirField;
    }

    TextArea getLogArea() {
        return logArea;
    }

    Label getStatusLabel() {
        return statusLabel;
    }

    void appendLog(String line) {
        logArea.appendText(line + "\n");
    }
}
