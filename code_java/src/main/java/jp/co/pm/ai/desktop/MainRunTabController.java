package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;

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
import javafx.scene.control.OverrunStyle;
import javafx.scene.control.ScrollBar;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.TextField;
import javafx.scene.control.Tooltip;
import javafx.scene.input.Clipboard;
import javafx.scene.input.ClipboardContent;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyCodeCombination;
import javafx.scene.input.KeyCombination;
import javafx.scene.effect.DropShadow;
import javafx.scene.input.KeyEvent;
import javafx.scene.layout.StackPane;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;
import javafx.util.StringConverter;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.AppVersionInfo;
import jp.co.pm.ai.desktop.config.PersonBadgeStyle;
import jp.co.pm.ai.desktop.io.DesktopFileOpener;
import jp.co.pm.ai.desktop.ui.PersonBadgeNodeFactory;

/** Run/log tab; layout in {@code MainRunTab.fxml}. */
public final class MainRunTabController {

    private static final int MAX_PERSISTED_LOG_LINES = 2000;

    private static final String DEFAULT_FONT_FAMILY_LABEL = "システム既定";

    private static final List<Double> PRESET_FONT_SIZES =
            List.of(9d, 10d, 11d, 12d, 13d, 14d, 15d, 16d, 18d, 20d, 22d, 24d);

    private MainShellController shell;

    @FXML
    private TextField workbookField;

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
    private StackPane stage1NetworkCacheBadgeHost;

    @FXML
    private Button stage1RunButton;

    @FXML
    private CheckBox stage2WriteExcelCheckBox;

    @FXML
    private CheckBox stage2SkipInProgressDispatchCheckBox;

    @FXML
    private ComboBox<String> stage2ResultBookFontCombo;

    @FXML
    private Button copyAllLogButton;

    @FXML
    private Button clearLogButton;

    @FXML
    private Label prismPipelineLabel;

    @FXML
    private Label appVersionLabel;

    @FXML
    private Label masterWorkbookOpenHintLabel;

    @FXML
    private Label summaryWorkbookOpenHintLabel;

    private final ObservableList<String> logLinesAll = FXCollections.observableArrayList();
    private final FilteredList<String> logLinesVisible =
            new FilteredList<>(logLinesAll, s -> true);

    private Font appliedLogFont = Font.getDefault();

    private final AtomicBoolean suppressLogFontEvents = new AtomicBoolean(false);

    private final AtomicBoolean suppressStage2ResultFontEvents = new AtomicBoolean(false);

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
        installStageRunButtonDepth(stage1RunButton, Color.rgb(14, 116, 144, 0.35));
        if (prismPipelineLabel != null) {
            prismPipelineLabel.setText(PrismGpuBootstrapStatus.runTabSummary());
        }
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
        if (stage2SkipInProgressDispatchCheckBox != null) {
            stage2SkipInProgressDispatchCheckBox
                    .selectedProperty()
                    .addListener(
                            (o, a, b) -> {
                                if (shell != null) {
                                    shell.scheduleDesktopSessionSave();
                                }
                            });
        }
        if (stage2ResultBookFontCombo != null) {
            List<String> stage2Families = new ArrayList<>();
            stage2Families.add(DEFAULT_FONT_FAMILY_LABEL);
            stage2Families.addAll(installed);
            stage2ResultBookFontCombo.getItems().setAll(stage2Families);
            stage2ResultBookFontCombo.getSelectionModel().selectFirst();
            stage2ResultBookFontCombo
                    .valueProperty()
                    .addListener(
                            (o, a, b) -> {
                                if (!suppressStage2ResultFontEvents.get()
                                        && shell != null) {
                                    shell.scheduleDesktopSessionSave();
                                }
                            });
        }
    }

    /** フラットボタン用のごく弱いドロップシャドウ（パルスなし）。 */
    private static void installStageRunButtonDepth(Button button, Color shadowColor) {
        if (button == null) {
            return;
        }
        DropShadow depth = new DropShadow();
        depth.setColor(shadowColor);
        depth.setRadius(10);
        depth.setSpread(0.12);
        depth.setOffsetY(2);
        button.setEffect(depth);
    }

    private void setupLogListView() {
        logListView.setItems(logLinesVisible);
        applyLogListFixedCellHeight();
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
                                setWrapText(false);
                                setTextOverrun(OverrunStyle.CLIP);
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

    /**
     * 可変行高（{@code setFixedCellSize(-1)}）と折り返しの組み合わせは VirtualFlow が極端なセル数を見積もり、
     * {@code index exceeds maxCellCount} やヒープ枯渇を招くことがある。フォントに応じた正の固定高で抑える。
     * ログは折り返さず右端でクリップするため、セル高は1行分に近い値とする。
     */
    private void applyLogListFixedCellHeight() {
        if (logListView == null) {
            return;
        }
        double lineHeight = appliedLogFont.getSize() * 1.35;
        double cell = Math.clamp(lineHeight * 1.45, 22.0, 52.0);
        logListView.setFixedCellSize(cell);
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
                        "選択をコピー (Ctrl+C)");
        copySelectedItem.setOnAction(e -> copySelectedLogLinesToClipboard());
        MenuItem copyAllItem =
                new MenuItem(
                        "全ログをコピー（バッファ全行）");
        copyAllItem.setOnAction(e -> copyAllBufferedLogToClipboard());
        MenuItem copyVisibleItem =
                new MenuItem(
                        "表示中のログをコピー");
        copyVisibleItem.setOnAction(e -> copyVisibleLogLinesToClipboard());
        logListView.setContextMenu(
                new ContextMenu(copySelectedItem, copyVisibleItem, copyAllItem));
    }

    /** Full buffer (ignores filter); same as toolbar 全ログをコピー. */
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
        refreshAppVersionLabel();
        refreshOpenWorkbookHintLabels();
    }

    /**
     * 実行・ログタブの「開く」横ラベルを環境変数（マスタ名・サマリブック）に合わせて更新する。
     */
    void refreshOpenWorkbookHintLabels() {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        if (masterWorkbookOpenHintLabel != null) {
            String alt = ui.getOrDefault(AppPaths.KEY_PM_AI_MASTER_WORKBOOK, "").trim();
            if (!alt.isEmpty()) {
                masterWorkbookOpenHintLabel.setText(Path.of(alt).getFileName().toString());
            } else {
                String mf = ui.getOrDefault(AppPaths.KEY_MASTER_WORKBOOK_FILE, "").trim();
                masterWorkbookOpenHintLabel.setText(mf.isEmpty() ? "master.xlsm" : mf);
            }
        }
        if (summaryWorkbookOpenHintLabel != null) {
            summaryWorkbookOpenHintLabel.setText(
                    AppPaths.summaryAiDispatchXlsxPath(ui).getFileName().toString());
        }
    }

    /** 実行・ログタブの {@code version.txt} 表示を更新する（ポータブル同期後など）。 */
    void refreshAppVersionLabel() {
        if (appVersionLabel == null || shell == null) {
            return;
        }
        Path cwd = Paths.get(System.getProperty("user.dir", "."));
        String v = AppVersionInfo.resolveDisplayedVersion(cwd, shell.snapshotUiEnv());
        appVersionLabel.setText("バージョン " + v);
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
            applyLogListFixedCellHeight();
            logListView.refresh();
        }
    }

    @FXML
    private void onStage1RunButtonAction() {
        shell.triggerStage1();
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
    private void onOpenMasterWorkbookAction() {
        if (shell == null) {
            return;
        }
        Path p =
                AppPaths.resolveMasterWorkbookPathForDesktopOpen(
                        shell.snapshotUiEnv(),
                        shell.effectiveTaskInputWorkbookPathForShell());
        if (!Files.isRegularFile(p)) {
            appendLog(
                    "[master-workbook] file not found: "
                            + p
                            + " (set "
                            + AppPaths.KEY_PM_AI_MASTER_WORKBOOK
                            + " / "
                            + AppPaths.KEY_MASTER_WORKBOOK_FILE
                            + ", or check "
                            + AppPaths.KEY_PM_AI_REPO_ROOT
                            + ")");
            return;
        }
        try {
            DesktopFileOpener.openFile(p);
            appendLog("[master-workbook] opened: " + p.toAbsolutePath().normalize());
        } catch (Exception e) {
            appendLog("[master-workbook] open failed: " + e.getMessage());
        }
    }

    @FXML
    private void onOpenDispatchUsageGuideDocxAction() {
        if (shell == null) {
            return;
        }
        Path p = AppPaths.resolveDispatchUsageGuideDocx(shell.snapshotUiEnv());
        if (!Files.isRegularFile(p)) {
            appendLog(
                    "[dispatch-usage-docx] file not found: "
                            + p
                            + " (expected "
                            + AppPaths.DISPATCH_USAGE_GUIDE_DOCX
                            + " under "
                            + AppPaths.KEY_PM_AI_REPO_ROOT
                            + ")");
            return;
        }
        try {
            DesktopFileOpener.openFile(p);
            appendLog("[dispatch-usage-docx] opened: " + p.toAbsolutePath().normalize());
        } catch (Exception e) {
            appendLog("[dispatch-usage-docx] open failed: " + e.getMessage());
        }
    }

    @FXML
    private void onOpenManualAction() {
        if (shell == null) {
            return;
        }
        Path p = AppPaths.resolveManualIndexHtml(shell.snapshotUiEnv());
        if (!Files.isRegularFile(p)) {
            appendLog(
                    "[manual] file not found: "
                            + p
                            + " (publish HTML per manual/README.md, or set "
                            + AppPaths.KEY_PM_AI_REPO_ROOT
                            + " if the repository root is wrong)");
            return;
        }
        try {
            DesktopFileOpener.openFile(p);
            appendLog("[manual] opened: " + p.toAbsolutePath().normalize());
        } catch (Exception e) {
            appendLog("[manual] open failed: " + e.getMessage());
        }
    }

    @FXML
    private void onOpenSummaryAiDispatchAction() {
        if (shell == null) {
            return;
        }
        Path p = AppPaths.summaryAiDispatchXlsxPath(shell.snapshotUiEnv());
        if (!Files.isRegularFile(p)) {
            appendLog(
                    "[summary-ai-dispatch] file not found: "
                            + p
                            + " (set "
                            + AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK
                            + " to open another book, or "
                            + AppPaths.KEY_PM_AI_REPO_ROOT
                            + " if the repository root is wrong)");
            return;
        }
        try {
            DesktopFileOpener.openFile(p);
            appendLog("[summary-ai-dispatch] opened: " + p.toAbsolutePath().normalize());
        } catch (Exception e) {
            appendLog("[summary-ai-dispatch] open failed: " + e.getMessage());
        }
    }

    @FXML
    private void onCopyAllLogButtonAction() {
        copyAllBufferedLogToClipboard();
    }

    @FXML
    private void onClearLogButtonAction() {
        clearMainRunTabLog();
    }

    /** メイン実行タブのログ一覧を空にする（クリアボタンと同一。ポータルバージョンアップ完了後など）。 */
    void clearMainRunTabLog() {
        logLinesAll.clear();
        if (shell != null) {
            shell.scheduleDesktopSessionSave();
        }
    }

    TextField getWorkbookField() {
        return workbookField;
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
     * 段階1／段階2 実行中は段階1ボタンの再実行を無効化する（進捗・中断はメインシェルツールバーのみ）。段階2実行ボタンは
     * {@link PlanInputTabController} 側。
     */
    void setStageRunProgressVisible(boolean stage1Running, boolean stage2Running) {
        boolean busy = stage1Running || stage2Running;
        if (stage1RunButton != null) {
            stage1RunButton.setDisable(busy);
        }
    }

    /**
     * ネットワークソースが使えずキャッシュを読んだとき、段階1ボタン左にバッジを表示する。
     */
    void setStage1NetworkCacheBadge(boolean visible, PersonBadgeStyle style, String labelText) {
        Platform.runLater(
                () -> {
                    if (stage1NetworkCacheBadgeHost == null) {
                        return;
                    }
                    stage1NetworkCacheBadgeHost.getChildren().clear();
                    stage1NetworkCacheBadgeHost.setManaged(visible);
                    stage1NetworkCacheBadgeHost.setVisible(visible);
                    if (!visible || style == null) {
                        return;
                    }
                    String t =
                            labelText != null && !labelText.isBlank() ? labelText.strip() : "キャッシュ";
                    StackPane graphic = PersonBadgeNodeFactory.createBadge(t, style, 1.0, 14.0);
                    Tooltip.install(
                            graphic,
                            new Tooltip(
                                    "PM_AI_TASK_INPUT_SOURCE_DIR または "
                                            + "PM_AI_ACTUAL_DETAIL_SOURCE_DIR "
                                            + "を参照できず、リポジトリ配下の最終キャッシュを使用して段階1／段階2に渡します。"));
                    stage1NetworkCacheBadgeHost.getChildren().add(graphic);
                });
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

    /** When {@code true}, stage-2 passes {@code PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH=1}. */
    boolean snapshotStage2SkipInProgressDispatch() {
        return stage2SkipInProgressDispatchCheckBox != null
                && stage2SkipInProgressDispatchCheckBox.isSelected();
    }

    void applyStage2SkipInProgressDispatchFromSession(boolean skipInProgress) {
        if (stage2SkipInProgressDispatchCheckBox != null) {
            stage2SkipInProgressDispatchCheckBox.setSelected(skipInProgress);
        }
    }

    String snapshotStage2ResultBookFont() {
        if (stage2ResultBookFontCombo == null) {
            return "";
        }
        String v = stage2ResultBookFontCombo.getValue();
        if (v == null
                || v.isBlank()
                || v.equals(DEFAULT_FONT_FAMILY_LABEL)) {
            return "";
        }
        return v.trim();
    }

    void applyStage2ResultBookFontFromSession(String family) {
        if (stage2ResultBookFontCombo == null) {
            return;
        }
        suppressStage2ResultFontEvents.set(true);
        try {
            if (family != null && !family.isBlank()) {
                if (!stage2ResultBookFontCombo.getItems().contains(family)) {
                    stage2ResultBookFontCombo.getItems().add(1, family);
                }
                stage2ResultBookFontCombo.setValue(family);
            } else {
                stage2ResultBookFontCombo.getSelectionModel().selectFirst();
            }
        } finally {
            suppressStage2ResultFontEvents.set(false);
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
