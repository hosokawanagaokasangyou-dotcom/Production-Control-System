package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;

import javafx.animation.Interpolator;
import javafx.animation.KeyFrame;
import javafx.animation.KeyValue;
import javafx.animation.Timeline;
import javafx.application.Platform;
import javafx.beans.binding.Bindings;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.collections.transformation.FilteredList;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ComboBox;
import javafx.scene.control.ContextMenu;
import javafx.scene.control.Label;
import javafx.scene.control.ListCell;
import javafx.scene.control.ListView;
import javafx.scene.control.MenuItem;
import javafx.scene.control.MultipleSelectionModel;
import javafx.scene.control.ScrollBar;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.TextField;
import javafx.scene.input.Clipboard;
import javafx.scene.input.ClipboardContent;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyCodeCombination;
import javafx.scene.input.KeyCombination;
import javafx.scene.effect.DropShadow;
import javafx.scene.input.KeyEvent;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;
import javafx.util.Duration;
import javafx.util.StringConverter;

import jp.co.pm.ai.desktop.io.DesktopFileOpener;

/** Run/log tab; layout in {@code MainRunTab.fxml}. */
public final class MainRunTabController {

    private static final int MAX_PERSISTED_LOG_LINES = 2000;

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
    private TextField stage2ProductionPlanField;

    @FXML
    private TextField stage2MemberScheduleField;

    @FXML
    private Button stage1RunButton;

    @FXML
    private Button stage2RunButton;

    @FXML
    private CheckBox stage2WriteExcelCheckBox;

    @FXML
    private Button copyAllLogButton;

    @FXML
    private Button clearLogButton;

    private final ObservableList<String> logLinesAll = FXCollections.observableArrayList();
    private final FilteredList<String> logLinesVisible =
            new FilteredList<>(logLinesAll, s -> true);

    private Font appliedLogFont = Font.getDefault();

    private final AtomicBoolean suppressLogFontEvents = new AtomicBoolean(false);

    private final AtomicBoolean suppressRunLogSessionPersistence = new AtomicBoolean(false);

    private double pendingSessionLogScroll = Double.NaN;

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
                            if (shell != null
                                    && !suppressRunLogSessionPersistence.get()) {
                                shell.scheduleDesktopSessionSave();
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
        installStageRunButtonGlow(stage1RunButton, Color.rgb(0, 229, 255, 0.72));
        installStageRunButtonGlow(stage2RunButton, Color.rgb(255, 171, 64, 0.78));
        if (stage2WriteExcelCheckBox != null) {
            stage2WriteExcelCheckBox
                    .selectedProperty()
                    .addListener(
                            (o, a, b) -> {
                                if (shell != null) {
                                    shell.scheduleDesktopSessionSave();
                                }
                            });
        }
    }

    /**
     * Soft pulsing outer glow (matches cyan / amber gradients in {@code pm-ai-desktop.css}).
     */
    private static void installStageRunButtonGlow(Button button, Color glowColor) {
        if (button == null) {
            return;
        }
        DropShadow glow = new DropShadow();
        glow.setColor(glowColor);
        glow.setRadius(20);
        glow.setSpread(0.42);
        button.setEffect(glow);

        Timeline pulse =
                new Timeline(
                        new KeyFrame(
                                Duration.ZERO,
                                new KeyValue(glow.radiusProperty(), 14, Interpolator.EASE_BOTH)),
                        new KeyFrame(
                                Duration.millis(1600),
                                new KeyValue(glow.radiusProperty(), 38, Interpolator.EASE_BOTH)));
        pulse.setAutoReverse(true);
        pulse.setCycleCount(Timeline.INDEFINITE);
        pulse.play();
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
    private void onStage1RunButtonAction() {
        shell.triggerStage1();
    }

    @FXML
    private void onStage2RunButtonAction() {
        shell.triggerStage2();
    }

    @FXML
    private void onOpenStage2ProductionPlanAction() {
        openExcelBesideField(stage2ProductionPlanField, "stage2-production-plan");
    }

    @FXML
    private void onOpenStage2MemberScheduleAction() {
        openExcelBesideField(stage2MemberScheduleField, "stage2-member-schedule");
    }

    @FXML
    private void onCopyAllLogButtonAction() {
        copyAllBufferedLogToClipboard();
    }

    @FXML
    private void onClearLogButtonAction() {
        logLinesAll.clear();
        if (shell != null) {
            shell.scheduleDesktopSessionSave();
        }
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

    /**
     * Fills read-only paths after stage-2 success (newest files under {@link
     * jp.co.pm.ai.desktop.config.AppPaths#defaultPlanningOutputDir}).
     */
    void setStage2ArtifactPaths(String productionPlanPath, String memberSchedulePath) {
        if (stage2ProductionPlanField != null) {
            stage2ProductionPlanField.setText(productionPlanPath != null ? productionPlanPath : "");
        }
        if (stage2MemberScheduleField != null) {
            stage2MemberScheduleField.setText(memberSchedulePath != null ? memberSchedulePath : "");
        }
        if (shell != null) {
            shell.scheduleDesktopSessionSave();
        }
    }

    String snapshotStage2ProductionPlanPath() {
        if (stage2ProductionPlanField == null || stage2ProductionPlanField.getText() == null) {
            return "";
        }
        return stage2ProductionPlanField.getText().trim();
    }

    String snapshotStage2MemberSchedulePath() {
        if (stage2MemberScheduleField == null || stage2MemberScheduleField.getText() == null) {
            return "";
        }
        return stage2MemberScheduleField.getText().trim();
    }

    /** When {@code true}, stage-2 passes {@code PM_AI_STAGE2_WRITE_EXCEL=1}; unchecked writes JSON only. */
    boolean snapshotStage2WriteExcel() {
        return stage2WriteExcelCheckBox == null || stage2WriteExcelCheckBox.isSelected();
    }

    void applyStage2WriteExcelFromSession(boolean writeExcel) {
        if (stage2WriteExcelCheckBox != null) {
            stage2WriteExcelCheckBox.setSelected(writeExcel);
        }
    }

    void appendLog(String line) {
        appendLog(line, true);
    }

    /**
     * Appends one log line. When {@code scrollToEnd} is false, vertical scroll is unchanged (used for boot
     * lines before restoring session scroll).
     */
    void appendLog(String line, boolean scrollToEnd) {
        Runnable add =
                () -> {
                    logLinesAll.add(line);
                    if (scrollToEnd) {
                        int n = logLinesVisible.size();
                        if (n > 0) {
                            logListView.scrollTo(n - 1);
                        }
                    }
                };
        if (Platform.isFxApplicationThread()) {
            add.run();
        } else {
            Platform.runLater(add);
        }
    }

    /**
     * Restores log buffer, filter, and schedules scroll restore after layout (see {@link
     * #flushPendingSessionLogScroll()}).
     */
    void restoreRunLogUiFromSession(
            String filterName, List<String> lines, double scrollProportion) {
        suppressRunLogSessionPersistence.set(true);
        try {
            if (lines != null && !lines.isEmpty()) {
                logLinesAll.setAll(lines);
            } else {
                logLinesAll.clear();
            }
            logFilterCombo.setValue(LogViewFilter.fromStoredName(filterName));
        } finally {
            suppressRunLogSessionPersistence.set(false);
        }
        pendingSessionLogScroll = scrollProportion;
    }

    /** Applies {@link #pendingSessionLogScroll} once the log {@link ListView} is laid out. */
    void flushPendingSessionLogScroll() {
        double p = pendingSessionLogScroll;
        pendingSessionLogScroll = Double.NaN;
        if (!Double.isFinite(p)) {
            return;
        }
        applyLogScrollProportion(p);
        Platform.runLater(() -> applyLogScrollProportion(p));
    }

    String snapshotLogFilterName() {
        LogViewFilter v = logFilterCombo != null ? logFilterCombo.getValue() : null;
        return v != null ? v.name() : LogViewFilter.ALL.name();
    }

    List<String> snapshotPersistedLogLines() {
        int n = logLinesAll.size();
        if (n <= MAX_PERSISTED_LOG_LINES) {
            return List.copyOf(logLinesAll);
        }
        return List.copyOf(logLinesAll.subList(n - MAX_PERSISTED_LOG_LINES, n));
    }

    double snapshotLogScrollProportion() {
        return readVerticalScrollProportion(logListView);
    }

    private static double readVerticalScrollProportion(ListView<?> listView) {
        if (listView == null) {
            return Double.NaN;
        }
        ScrollBar sb = (ScrollBar) listView.lookup(".scroll-bar:vertical");
        if (sb == null) {
            return Double.NaN;
        }
        double min = sb.getMin();
        double max = sb.getMax();
        double v = sb.getValue();
        if (max <= min) {
            return 0d;
        }
        return (v - min) / (max - min);
    }

    private void applyLogScrollProportion(double proportion) {
        if (logListView == null || !Double.isFinite(proportion)) {
            return;
        }
        double p = Math.max(0d, Math.min(1d, proportion));
        ScrollBar sb = (ScrollBar) logListView.lookup(".scroll-bar:vertical");
        if (sb == null) {
            return;
        }
        double min = sb.getMin();
        double max = sb.getMax();
        if (max > min) {
            sb.setValue(min + p * (max - min));
        }
    }

    private void openExcelBesideField(TextField field, String logTag) {
        String raw = field != null && field.getText() != null ? field.getText().trim() : "";
        if (raw.isEmpty()) {
            appendLog("[" + logTag + "] path is empty");
            return;
        }
        Path p = Paths.get(raw);
        if (!Files.isRegularFile(p)) {
            appendLog("[" + logTag + "] file not found: " + p);
            return;
        }
        try {
            DesktopFileOpener.openFile(p);
            appendLog("[" + logTag + "] opened: " + p.toAbsolutePath().normalize());
        } catch (Exception e) {
            appendLog("[" + logTag + "] open failed: " + e.getMessage());
        }
    }
}
