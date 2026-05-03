package jp.co.pm.ai.desktop;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;

import javafx.application.Platform;
import javafx.beans.binding.Bindings;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.collections.transformation.FilteredList;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.ContextMenu;
import javafx.scene.control.Label;
import javafx.scene.control.ListCell;
import javafx.scene.control.ListView;
import javafx.scene.control.MenuItem;
import javafx.scene.control.MultipleSelectionModel;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.TextField;
import javafx.scene.input.Clipboard;
import javafx.scene.input.ClipboardContent;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyCodeCombination;
import javafx.scene.input.KeyCombination;
import javafx.scene.input.KeyEvent;
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
    private ListView<String> logListView;

    @FXML
    private ComboBox<LogViewFilter> logFilterCombo;

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

    @FXML
    private Button copyAllLogButton;

    @FXML
    private Button clearLogButton;

    private final ObservableList<String> logLinesAll = FXCollections.observableArrayList();
    private final FilteredList<String> logLinesVisible =
            new FilteredList<>(logLinesAll, s -> true);

    private Font appliedLogFont = Font.getDefault();

    private final AtomicBoolean suppressLogFontEvents = new AtomicBoolean(false);

    @FXML
    private void initialize() {
        logFilterCombo.getItems().setAll(LogViewFilter.values());
        logFilterCombo.setValue(LogViewFilter.ALL);
        logFilterCombo.setConverter(
                new StringConverter<>() {
                    @Override
                    public String toString(LogViewFilter f) {
                        return f != null ? f.getLabel() : "";
                    }

                    @Override
                    public LogViewFilter fromString(String string) {
                        if (string == null || string.isBlank()) {
                            return LogViewFilter.ALL;
                        }
                        for (LogViewFilter v : LogViewFilter.values()) {
                            if (v.getLabel().equals(string)) {
                                return v;
                            }
                        }
                        return LogViewFilter.ALL;
                    }
                });
        logFilterCombo
                .valueProperty()
                .addListener(
                        (o, a, b) -> {
                            if (b != null) {
                                logLinesVisible.setPredicate(b::test);
                            }
                        });

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

        setupLogListView();
        if (copyAllLogButton != null) {
            copyAllLogButton.disableProperty().bind(Bindings.isEmpty(logLinesAll));
        }
        if (clearLogButton != null) {
            clearLogButton.disableProperty().bind(Bindings.isEmpty(logLinesAll));
        }
        applyLogAreaFont();
    }

    private void setupLogListView() {
        logListView.setItems(logLinesVisible);
        logListView.setFixedCellSize(-1);
        logListView.setFocusTraversable(true);
        logListView.setCellFactory(
                lv ->
                        new ListCell<>() {
                            @Override
                            protected void updateItem(String item, boolean empty) {
                                super.updateItem(item, empty);
                                getStyleClass()
                                        .removeAll(
                                                "log-cell",
                                                "log-kind-error",
                                                "log-kind-warn",
                                                "log-dark");
                                if (empty || item == null) {
                                    setText(null);
                                    setGraphic(null);
                                    return;
                                }
                                setText(item);
                                setWrapText(true);
                                setFont(appliedLogFont);
                                double w = logListView.getWidth() - 28;
                                if (w > 0) {
                                    setMaxWidth(w);
                                }
                                getStyleClass().add("log-cell");
                                if (shell != null && shell.currentDesktopTheme().isDarkUi()) {
                                    getStyleClass().add("log-dark");
                                }
                                switch (LogLineKind.classify(item)) {
                                    case ERROR -> getStyleClass().add("log-kind-error");
                                    case WARN -> getStyleClass().add("log-kind-warn");
                                    case NORMAL -> {
                                        /* default row chrome only */
                                    }
                                }
                            }
                        });
        logListView
                .widthProperty()
                .addListener(
                        (o, a, b) -> {
                            if (logListView != null) {
                                logListView.refresh();
                            }
                        });
        logListView.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        installLogClipboardSupport();
    }

    private void installLogClipboardSupport() {
        var copyKeys = new KeyCodeCombination(KeyCode.C, KeyCombination.SHORTCUT_DOWN);
        var selectAllKeys = new KeyCodeCombination(KeyCode.A, KeyCombination.SHORTCUT_DOWN);
        logListView.addEventFilter(
                KeyEvent.KEY_PRESSED,
                e -> {
                    if (copyKeys.match(e)) {
                        copySelectedLogLinesToClipboard();
                        e.consume();
                    } else if (selectAllKeys.match(e)) {
                        logListView.getSelectionModel().selectAll();
                        e.consume();
                    }
                });
        MenuItem copySelectedItem =
                new MenuItem(
                        "\u9078\u629e\u3092\u30b3\u30d4\u30fc (Ctrl+C)");
        copySelectedItem.setOnAction(e -> copySelectedLogLinesToClipboard());
        MenuItem copyAllItem =
                new MenuItem(
                        "\u5168\u30ed\u30b0\u3092\u30b3\u30d4\u30fc\uFF08\u30d0\u30c3\u30d5\u30a1\u5168\u884c\uFF09");
        copyAllItem.setOnAction(e -> copyAllBufferedLogToClipboard());
        MenuItem copyVisibleItem =
                new MenuItem(
                        "\u8868\u793a\u4e2d\u306e\u30ed\u30b0\u3092\u30b3\u30d4\u30fc");
        copyVisibleItem.setOnAction(e -> copyVisibleLogLinesToClipboard());
        logListView.setContextMenu(
                new ContextMenu(copySelectedItem, copyVisibleItem, copyAllItem));
    }

    /** Full buffer (ignores filter); same as toolbar \u5168\u30ed\u30b0\u3092\u30b3\u30d4\u30fc. */
    private void copyAllBufferedLogToClipboard() {
        if (logLinesAll.isEmpty()) {
            return;
        }
        String text = String.join("\n", logLinesAll);
        ClipboardContent cc = new ClipboardContent();
        cc.putString(text);
        Clipboard.getSystemClipboard().setContent(cc);
    }

    private void copyVisibleLogLinesToClipboard() {
        if (logLinesVisible.isEmpty()) {
            return;
        }
        String text = String.join("\n", logLinesVisible);
        ClipboardContent cc = new ClipboardContent();
        cc.putString(text);
        Clipboard.getSystemClipboard().setContent(cc);
    }

    private void copySelectedLogLinesToClipboard() {
        ObservableList<String> visible = logListView.getItems();
        MultipleSelectionModel<String> sm = logListView.getSelectionModel();
        ArrayList<Integer> indices = new ArrayList<>(sm.getSelectedIndices());
        if (indices.isEmpty()) {
            int fi = logListView.getFocusModel().getFocusedIndex();
            if (fi >= 0) {
                indices.add(fi);
            }
        }
        if (indices.isEmpty()) {
            return;
        }
        Collections.sort(indices);
        StringBuilder sb = new StringBuilder();
        for (int i : indices) {
            if (i >= 0 && i < visible.size()) {
                if (sb.length() > 0) {
                    sb.append('\n');
                }
                sb.append(visible.get(i));
            }
        }
        if (sb.length() == 0) {
            return;
        }
        ClipboardContent cc = new ClipboardContent();
        cc.putString(sb.toString());
        Clipboard.getSystemClipboard().setContent(cc);
    }

    /** Reapply row styles when UI theme (dark/light) changes. */
    void refreshLogThemeCells() {
        if (logListView != null) {
            logListView.refresh();
        }
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
        if (logFontFamilyCombo == null || logFontSizeCombo == null) {
            return;
        }
        String choice = logFontFamilyCombo.getValue();
        Double szObj = logFontSizeCombo.getValue();
        double size =
                szObj != null && szObj > 0 && Double.isFinite(szObj)
                        ? szObj
                        : Font.getDefault().getSize();
        if (choice == null || choice.equals(DEFAULT_FONT_FAMILY_LABEL)) {
            appliedLogFont = Font.font(size);
        } else {
            appliedLogFont = Font.font(choice, size);
        }
        if (logListView != null) {
            logListView.refresh();
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

    @FXML
    private void onCopyAllLogButtonAction() {
        copyAllBufferedLogToClipboard();
    }

    @FXML
    private void onClearLogButtonAction() {
        logLinesAll.clear();
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

    ListView<String> getLogListView() {
        return logListView;
    }

    Label getStatusLabel() {
        return statusLabel;
    }

    void appendLog(String line) {
        Runnable add =
                () -> {
                    logLinesAll.add(line);
                    int n = logLinesVisible.size();
                    if (n > 0) {
                        logListView.scrollTo(n - 1);
                    }
                };
        if (Platform.isFxApplicationThread()) {
            add.run();
        } else {
            Platform.runLater(add);
        }
    }
}
