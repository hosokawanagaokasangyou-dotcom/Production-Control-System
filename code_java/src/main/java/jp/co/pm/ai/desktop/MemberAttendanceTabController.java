package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.YearMonth;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicLong;
import java.util.function.Consumer;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import javafx.animation.KeyFrame;
import javafx.animation.PauseTransition;
import javafx.animation.Timeline;
import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonBar;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.Tooltip;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;
import javafx.util.Duration;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.dispatch.AttendanceOvertimePreview;
import jp.co.pm.ai.desktop.ui.AttendanceGridCellSizing;
import jp.co.pm.ai.desktop.ui.AttendanceSyncStatusPane;
import jp.co.pm.ai.desktop.ui.ButtonAttentionGlow;
import jp.co.pm.ai.desktop.ui.EditableMemberAttendanceGridPane;
import jp.co.pm.ai.desktop.ui.FiscalYearPeriod;
import jp.co.pm.ai.desktop.ui.FourDigitConfirmationDialog;
import jp.co.pm.ai.desktop.ui.InlineMonthCalendarPane;
import jp.co.pm.ai.desktop.ui.MemberAttendanceMemberEditDialog;
import jp.co.pm.ai.desktop.ui.MemberHourlyAttendanceDialog;

/** メンバー勤怠（カレンダー方式）編集タブ。 */
public class MemberAttendanceTabController {

    private static final ObjectMapper JSON = new ObjectMapper();
    private static final int MEMBER_GRID_CACHE_MAX_MONTHS = 12;
    private static final Duration MEMBER_GRID_IDLE_PREFETCH_DELAY = Duration.millis(650);

    @FXML
    private VBox gridHost;

    @FXML
    private VBox statusHost;

    @FXML
    private VBox monthCalendarHost;

    @FXML
    private Label statusLabel;

    @FXML
    private Label setupHintLabel;

    @FXML
    private Button saveButton;

    @FXML
    private Button setupButton;

    @FXML
    private Button restoreButton;

    @FXML
    private Button refreshButton;

    @FXML
    private Button openCalendarButton;

    @FXML
    private Button addMemberButton;

    @FXML
    private Button editMemberButton;

    @FXML
    private Button removeMemberButton;

    @FXML
    private Spinner<Integer> cellSizeSpinner;

    private MainShellController shell;
    private EditableMemberAttendanceGridPane gridPane;
    private AttendanceSyncStatusPane syncStatusPane;
    private InlineMonthCalendarPane monthCalendar;
    private ButtonAttentionGlow saveButtonGlow;
    private ButtonAttentionGlow setupButtonGlow;
    private final AtomicLong loadGeneration = new AtomicLong(0);
    private final PauseTransition gridReloadDebounce = new PauseTransition(Duration.millis(350));
    private final PauseTransition memberGridIdlePrefetch =
            new PauseTransition(MEMBER_GRID_IDLE_PREFETCH_DELAY);
    private final LinkedHashMap<String, JsonNode> memberGridCache =
            new LinkedHashMap<>(16, 0.75f, true) {
                @Override
                protected boolean removeEldestEntry(Map.Entry<String, JsonNode> eldest) {
                    return size() > MEMBER_GRID_CACHE_MAX_MONTHS;
                }
            };
    private volatile boolean memberGridPrefetchInFlight = false;
    private boolean attendanceLoadEnabled = false;
    private boolean suppressMonthGuard = false;
    private int tabProcessingDepth = 0;
    private int setupWizardGridOverlayDepth = 0;
    private String activeLoadingMessage = "処理中";
    private ProgressIndicator statusProgress;
    private Timeline statusActivityTick;
    private long statusActivityStartMs = 0L;

    public enum UnsavedPromptResult {
        SAVED,
        DISCARDED,
        CANCELLED
    }

    public void bindShell(MainShellController shell) {
        this.shell = shell;
        LocalDate today = LocalDate.now();
        if (statusHost != null && syncStatusPane == null) {
            syncStatusPane = new AttendanceSyncStatusPane();
            statusHost.getChildren().add(syncStatusPane);
        }
        if (gridHost != null && gridPane == null) {
            gridPane = new EditableMemberAttendanceGridPane();
            gridPane.setCellDetailHandler(
                    req -> {
                        Stage owner =
                                shell != null ? shell.primaryStageForDialogs() : null;
                        MemberHourlyAttendanceDialog.show(
                                owner,
                                req.member(),
                                req.date().toString(),
                                req.state().hourly(),
                                hourly ->
                                        gridPane.applyHourlyEdit(
                                                req.date(),
                                                req.member(),
                                                hourly,
                                                req.state().dayPreset()));
                    });
            gridPane.setDirtyListener(this::applyGridDirtyState);
            gridPane.setCommentDialogOwner(shell.primaryStageForDialogs());
            VBox.setVgrow(gridPane, Priority.ALWAYS);
            gridPane.setMaxHeight(Double.MAX_VALUE);
            gridHost.setMaxHeight(Double.MAX_VALUE);
            VBox.setVgrow(gridHost, Priority.ALWAYS);
            gridHost.getChildren().add(gridPane);
        }
        if (saveButton != null && saveButtonGlow == null) {
            saveButtonGlow = new ButtonAttentionGlow(saveButton);
        }
        if (setupButton != null && setupButtonGlow == null) {
            setupButtonGlow = new ButtonAttentionGlow(setupButton);
        }
        installGridCellSizeSpinner();
        applyGridCellSizeToPane(shell.attendanceGridCellSizePx());
        if (setupHintLabel != null) {
            setupHintLabel.getStyleClass().add("pm-member-attendance-setup-hint");
        }
        installToolbarTooltips();
        installStatusActivityRow();
        installMonthCalendar(today);
        installMemberGridIdlePrefetch();
    }

    private void installMemberGridIdlePrefetch() {
        memberGridIdlePrefetch.setOnFinished(e -> runMemberGridIdlePrefetchStep());
    }

    private void installStatusActivityRow() {
        if (statusLabel == null || statusLabel.getParent() == null) {
            return;
        }
        statusProgress = new ProgressIndicator();
        statusProgress.setPrefSize(16, 16);
        statusProgress.setMaxSize(16, 16);
        statusProgress.setVisible(false);
        HBox row = new HBox(8, statusProgress, statusLabel);
        row.setAlignment(Pos.CENTER_LEFT);
        if (statusLabel.getParent() instanceof VBox parent) {
            int idx = parent.getChildren().indexOf(statusLabel);
            if (idx >= 0) {
                parent.getChildren().set(idx, row);
            }
        }
        statusActivityTick =
                new Timeline(
                        new KeyFrame(
                                Duration.millis(400),
                                e -> updateStatusActivityLabel()));
        statusActivityTick.setCycleCount(Timeline.INDEFINITE);
    }

    private void installToolbarTooltips() {
        installTooltip(saveButton, "編集内容を JSON に保存し 勤怠カレンダー.xlsx を出力します");
        installTooltip(
                setupButton,
                "祝日・週末公休の取得とメンバー勤怠の会社カレンダー同期");
        installTooltip(restoreButton, "attendance-data.json の過去リビジョンから復元します");
        installTooltip(
                openCalendarButton,
                "勤怠カレンダー.xlsx を Excel で読み取り専用で開きます（未出力の場合は先に保存してください）");
        installTooltip(
                refreshButton,
                "JSON 正本から再読込します（未保存の変更がある場合は確認します）");
        installTooltip(addMemberButton, "名簿にメンバーを追加します（保存で JSON に反映）");
        installTooltip(editMemberButton, "選択したメンバー行の氏名・主担当を編集します");
        installTooltip(
                removeMemberButton,
                "選択したメンバーを名簿から削除します（氏名列をクリックして選択）");
        if (cellSizeSpinner != null) {
            installTooltip(cellSizeSpinner, "グリッドセルサイズ（会社カレンダーと共通）");
        }
    }

    private static void installTooltip(javafx.scene.control.Control control, String text) {
        if (control != null) {
            control.setTooltip(new Tooltip(text));
        }
    }

    private void installMonthCalendar(LocalDate today) {
        if (monthCalendarHost == null || monthCalendar != null) {
            return;
        }
        monthCalendar = new InlineMonthCalendarPane(true);
        monthCalendar.setSelectedDate(today);
        monthCalendarHost.getChildren().add(monthCalendar);
        gridReloadDebounce.setOnFinished(
                e -> {
                    if (attendanceLoadEnabled) {
                        loadGridFromPython();
                    }
                });
        monthCalendar
                .selectedDateProperty()
                .addListener(
                        (obs, oldDate, newDate) -> {
                            if (!attendanceLoadEnabled || newDate == null) {
                                return;
                            }
                            if (oldDate != null
                                    && YearMonth.from(oldDate).equals(YearMonth.from(newDate))) {
                                return;
                            }
                            if (suppressMonthGuard) {
                                scheduleGridReload();
                                return;
                            }
                            handleUnsavedThen(
                                    "表示月を変える",
                                    () -> scheduleGridReload(),
                                    () -> {
                                        suppressMonthGuard = true;
                                        monthCalendar.setSelectedDate(oldDate);
                                        suppressMonthGuard = false;
                                    });
                        });
    }

    private void handleUnsavedThen(
            String actionDescription, Runnable onProceed, Runnable onCancel) {
        if (!hasUnsavedEdits()) {
            onProceed.run();
            return;
        }
        UnsavedPromptResult result = promptUnsavedChanges(actionDescription);
        if (result == UnsavedPromptResult.CANCELLED) {
            onCancel.run();
            return;
        }
        if (result == UnsavedPromptResult.DISCARDED) {
            onProceed.run();
            return;
        }
        saveEditsAsync(
                saved -> {
                    if (saved) {
                        onProceed.run();
                    } else {
                        onCancel.run();
                    }
                });
    }

    private LocalDate selectedCalendarDate() {
        if (monthCalendar != null && monthCalendar.getSelectedDate() != null) {
            return monthCalendar.getSelectedDate();
        }
        return LocalDate.now();
    }

    public boolean hasUnsavedEdits() {
        return gridPane != null && gridPane.hasUnsavedEdits();
    }

    public void discardUnsavedEdits() {
        loadGridFromPython();
    }

    public void clearUnsavedWithoutReload() {
        if (gridPane != null) {
            gridPane.clearUnsavedEditFlags();
            applyGridDirtyState(false);
        }
    }

    public UnsavedPromptResult promptUnsavedChanges(String actionDescription) {
        if (!hasUnsavedEdits()) {
            return UnsavedPromptResult.DISCARDED;
        }
        Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
        if (shell != null) {
            alert.initOwner(shell.primaryStageForDialogs());
            shell.applyAlertStylesheets(alert);
        }
        alert.setTitle("未保存の変更");
        alert.setHeaderText(null);
        alert.setContentText(
                "メンバー勤怠に未保存の変更があります。"
                        + actionDescription
                        + "前に保存しますか？");
        ButtonType save = new ButtonType("保存", ButtonBar.ButtonData.OK_DONE);
        ButtonType discard = new ButtonType("保存しない", ButtonBar.ButtonData.NO);
        alert.getButtonTypes().setAll(save, discard, ButtonType.CANCEL);
        Optional<ButtonType> ans = alert.showAndWait();
        if (ans.isEmpty() || ans.get() == ButtonType.CANCEL) {
            return UnsavedPromptResult.CANCELLED;
        }
        if (ans.get() == discard) {
            return UnsavedPromptResult.DISCARDED;
        }
        return UnsavedPromptResult.SAVED;
    }

    public void saveEditsAsync(Consumer<Boolean> onComplete) {
        if (shell == null || gridPane == null) {
            if (onComplete != null) {
                onComplete.accept(false);
            }
            return;
        }
        if (!confirmSaveWithFourDigit()) {
            if (onComplete != null) {
                onComplete.accept(false);
            }
            return;
        }
        try {
            Map<String, Object> patch = gridPane.exportPatchJson();
            String json = JSON.writeValueAsString(patch);
            Path tmp = Files.createTempFile("pm-ai-member-attendance-", ".json");
            Files.writeString(tmp, json);
            runAsync(
                    shell.buildAttendanceDataIoRequest(
                            "merge_member_attendance", "--patch-file", tmp.toString()),
                    mergeNode -> {
                        gridPane.clearUnsavedEditFlags();
                        applyGridDirtyState(false);
                        clearMemberGridCache();
                        statusLabel.setText(
                                "保存・勤怠カレンダー.xlsx 出力完了: "
                                        + mergeNode.path("applied").asInt(0)
                                        + " セル → "
                                        + mergeNode.path("json_path").asText("")
                                        + " / "
                                        + mergeNode.path("calendar_xlsx_path").asText("")
                                        + " / シート "
                                        + mergeNode.path("sheets_updated").toString());
                        shell.refreshAttendanceReadiness();
                        refreshLocalReadiness();
                        if (onComplete != null) {
                            onComplete.accept(true);
                        }
                    },
                    false,
                    tmp,
                    null,
                    mergeSuccess -> {
                        if (!mergeSuccess && onComplete != null) {
                            onComplete.accept(false);
                        }
                    },
                    "メンバー勤怠を保存・Excel 出力中");
        } catch (Exception e) {
            statusLabel.setText("エラー: " + e.getMessage());
            if (onComplete != null) {
                onComplete.accept(false);
            }
        }
    }

    private void applyGridDirtyState(boolean dirty) {
        if (saveButtonGlow != null) {
            if (dirty) {
                saveButtonGlow.startIfIdle();
            } else {
                saveButtonGlow.stop();
            }
        }
        if (shell != null) {
            shell.onMemberAttendanceDirtyChanged(dirty);
        }
    }

    private void applySetupButtonAttention(JsonNode node) {
        if (setupButtonGlow == null) {
            return;
        }
        if (AttendanceSyncStatusPane.needsSetupAttention(node)) {
            setupButtonGlow.ensureActive();
        } else {
            setupButtonGlow.stop();
        }
    }

    private void scheduleGridReload() {
        gridReloadDebounce.playFromStart();
    }

    private static String memberGridCacheKey(int year, int month) {
        return year + "-" + month;
    }

    private void clearMemberGridCache() {
        memberGridCache.clear();
    }

    private void storeMemberGridCache(JsonNode node) {
        if (node == null || !node.path("ok").asBoolean(false)) {
            return;
        }
        int y = node.path("year").asInt(0);
        int m = node.path("month").asInt(0);
        if (y <= 0 || m <= 0) {
            return;
        }
        memberGridCache.put(memberGridCacheKey(y, m), node);
        prefetchAdjacentMemberGrids(y, m);
    }

    private void scheduleMemberGridIdlePrefetch() {
        if (!attendanceLoadEnabled || shell == null) {
            return;
        }
        memberGridIdlePrefetch.playFromStart();
    }

    private void runMemberGridIdlePrefetchStep() {
        if (!attendanceLoadEnabled
                || shell == null
                || tabProcessingDepth > 0
                || setupWizardGridOverlayDepth > 0
                || memberGridPrefetchInFlight) {
            scheduleMemberGridIdlePrefetch();
            return;
        }
        YearMonth center = YearMonth.from(selectedCalendarDate());
        YearMonth target = nextMemberGridPrefetchTarget(center);
        if (target == null) {
            return;
        }
        memberGridPrefetchInFlight = true;
        prefetchMemberGridQuiet(
                target,
                () -> {
                    memberGridPrefetchInFlight = false;
                    scheduleMemberGridIdlePrefetch();
                });
    }

    private YearMonth nextMemberGridPrefetchTarget(YearMonth center) {
        for (YearMonth ym : memberGridPrefetchMonths(center)) {
            String key = memberGridCacheKey(ym.getYear(), ym.getMonthValue());
            if (!memberGridCache.containsKey(key)) {
                return ym;
            }
        }
        return null;
    }

    private List<YearMonth> memberGridPrefetchMonths(YearMonth center) {
        if (shell == null) {
            return List.of();
        }
        FiscalYearPeriod period = shell.attendanceFiscalPeriod();
        LocalDate anchor = selectedCalendarDate();
        int fiscalYear = FiscalYearPeriod.fiscalYearLabelFor(anchor, period);
        List<YearMonth> candidates = new ArrayList<>(period.monthsInOrder(fiscalYear));
        candidates.sort(
                Comparator.comparingInt(ym -> memberGridMonthDistance(center, ym)));
        return candidates;
    }

    private static int memberGridMonthDistance(YearMonth center, YearMonth other) {
        return Math.abs(center.getYear() * 12 + center.getMonthValue()
                - other.getYear() * 12 - other.getMonthValue());
    }

    private void prefetchAdjacentMemberGrids(int year, int month) {
        if (!attendanceLoadEnabled || shell == null) {
            return;
        }
        YearMonth center = YearMonth.of(year, month);
        prefetchMemberGridQuiet(center.minusMonths(1));
        prefetchMemberGridQuiet(center.plusMonths(1));
    }

    private void prefetchMemberGridQuiet(YearMonth ym) {
        prefetchMemberGridQuiet(ym, null);
    }

    private void prefetchMemberGridQuiet(YearMonth ym, Runnable onComplete) {
        if (shell == null) {
            if (onComplete != null) {
                onComplete.run();
            }
            return;
        }
        String key = memberGridCacheKey(ym.getYear(), ym.getMonthValue());
        if (memberGridCache.containsKey(key)) {
            if (onComplete != null) {
                onComplete.run();
            }
            return;
        }
        PythonProcessRunner.runCaptureAsync(
                        shell.buildAttendanceDataIoRequest(
                                "member_grid",
                                Integer.toString(ym.getYear()),
                                Integer.toString(ym.getMonthValue())))
                .whenComplete(
                        (cap, err) ->
                                Platform.runLater(
                                        () -> {
                                            try {
                                                if (err == null
                                                        && cap != null
                                                        && cap.exitCode() == 0) {
                                                    JsonNode node =
                                                            JSON.readTree(
                                                                    AttendanceOvertimePreview
                                                                            .MasterReadSummaryJson
                                                                            .extractLastJsonLine(
                                                                                    cap.stdout()));
                                                    if (node.path("ok").asBoolean(false)) {
                                                        memberGridCache.put(key, node);
                                                    }
                                                }
                                            } catch (Exception ignored) {
                                                // ignore prefetch failures
                                            } finally {
                                                if (onComplete != null) {
                                                    onComplete.run();
                                                }
                                            }
                                        }));
    }

    private void applyMemberGridNode(JsonNode node, int year, int month) {
        if (gridPane != null) {
            gridPane.loadFromMemberGridJson(node);
        }
        statusLabel.setText(
                "読込 "
                        + year
                        + "/"
                        + month
                        + " メンバー="
                        + node.path("members").size()
                        + " revision="
                        + node.path("member_attendance_revision").asInt(0));
        scheduleMemberGridIdlePrefetch();
    }

    /** セッション・環境変数復元後に MainShell から呼ぶ。JSON 正本を読み込む。 */
    public void enableAttendanceLoadAndRefresh() {
        if (attendanceLoadEnabled) {
            loadGridFromPython();
            refreshLocalReadiness();
            return;
        }
        attendanceLoadEnabled = true;
        loadGridFromPython();
        refreshLocalReadiness();
    }

    /** メインシェルで当該タブが選択されたときに初回読込する。 */
    void onMainShellTabSelected() {
        if (!attendanceLoadEnabled) {
            enableAttendanceLoadAndRefresh();
        }
    }

    /** 環境変数・パス確定後の再読込（起動時・工場ワークスペース復元後）。 */
    public void reloadAttendanceDataFromJson() {
        if (!attendanceLoadEnabled) {
            enableAttendanceLoadAndRefresh();
            return;
        }
        clearMemberGridCache();
        loadGridFromPython();
        refreshLocalReadiness();
    }

    /** タブ表示済みのときだけ JSON を再読込する（起動時の一括読込は避ける）。 */
    void reloadAttendanceDataFromJsonIfEnabled() {
        if (!attendanceLoadEnabled) {
            return;
        }
        clearMemberGridCache();
        loadGridFromPython();
        refreshLocalReadiness();
    }

    /** 起動後バックグラウンド読込（MainShell コーディネータから呼ぶ）。 */
    void preloadInBackground(Consumer<Boolean> onComplete) {
        if (!attendanceLoadEnabled) {
            attendanceLoadEnabled = true;
        }
        loadGridFromPython(onComplete);
        refreshLocalReadiness();
    }

    private void installGridCellSizeSpinner() {
        if (cellSizeSpinner == null || shell == null) {
            return;
        }
        cellSizeSpinner.setValueFactory(
                new SpinnerValueFactory.IntegerSpinnerValueFactory(
                        AttendanceGridCellSizing.MIN_PX,
                        AttendanceGridCellSizing.MAX_PX,
                        shell.attendanceGridCellSizePx()));
        cellSizeSpinner
                .valueProperty()
                .addListener(
                        (obs, o, n) -> {
                            if (n != null && shell != null) {
                                shell.setAttendanceGridCellSizePx(n);
                            }
                        });
    }

    public void applyGridCellSizeToPane(int px) {
        if (gridPane != null) {
            gridPane.setCellSizePx(px);
        }
    }

    public void refreshRowHoverDimming() {
        if (gridPane != null) {
            gridPane.refreshRowHoverDimming();
        }
    }

    public void syncGridCellSizeSpinner(int px) {
        if (cellSizeSpinner == null) {
            return;
        }
        SpinnerValueFactory<Integer> vf = cellSizeSpinner.getValueFactory();
        if (vf instanceof SpinnerValueFactory.IntegerSpinnerValueFactory intVf
                && intVf.getValue() != px) {
            intVf.setValue(px);
        }
    }

    @FXML
    private void onSetupWizard() {
        if (shell != null) {
            beginSetupWizardGridOverlay();
            LocalDate selected = selectedCalendarDate();
            FiscalYearPeriod period = shell.attendanceFiscalPeriod();
            int fiscalYear =
                    FiscalYearPeriod.fiscalYearLabelFor(selected, period);
            AttendanceSetupWizard.show(
                    shell,
                    fiscalYear,
                    period,
                    selected.getYear(),
                    selected.getMonthValue(),
                    ok -> {
                        endSetupWizardGridOverlay();
                        if (ok) {
                            clearMemberGridCache();
                            loadGridFromPython();
                            shell.refreshAttendanceReadiness();
                            refreshLocalReadiness();
                        }
                    });
            endSetupWizardGridOverlay();
        }
    }

    @FXML
    private void onRestoreFromHistory() {
        if (shell == null) {
            return;
        }
        AttendanceJsonHistoryDialog.show(
                shell,
                () -> {
                    clearMemberGridCache();
                    loadGridFromPython();
                    refreshLocalReadiness();
                });
    }

    private void refreshLocalReadiness() {
        if (shell == null) {
            return;
        }
        LocalDate selected = selectedCalendarDate();
        shell.runAttendanceDataIoAsync(
                shell.buildAttendanceDataIoRequest(
                        "readiness",
                        Integer.toString(selected.getYear()),
                        Integer.toString(selected.getMonthValue())),
                node -> {
                    if (syncStatusPane != null) {
                        syncStatusPane.updateFromReadiness(node);
                    }
                    applySetupButtonAttention(node);
                },
                err -> {
                    if (statusLabel != null) {
                        statusLabel.setText("状態取得失敗: " + err);
                    }
                });
    }

    @FXML
    private void onSave() {
        saveEditsAsync(null);
    }

    @FXML
    private void onAddMember() {
        if (gridPane == null || shell == null) {
            return;
        }
        MemberAttendanceMemberEditDialog.showAdd(shell.primaryStageForDialogs())
                .ifPresent(
                        r -> {
                            clearMemberGridCache();
                            gridPane.addMember(r.name(), r.primaryRole());
                        });
    }

    @FXML
    private void onEditMember() {
        if (gridPane == null || shell == null) {
            return;
        }
        String selected = gridPane.selectedMemberName();
        if (selected == null || selected.isBlank()) {
            shell.showWarningDialog("メンバー編集", "編集するメンバー行（氏名列）をクリックして選択してください。");
            return;
        }
        MemberAttendanceMemberEditDialog.showEdit(
                        shell.primaryStageForDialogs(),
                        selected,
                        gridPane.primaryRoleFor(selected))
                .ifPresent(
                        r -> {
                            clearMemberGridCache();
                            gridPane.updateMember(
                                    selected, r.name(), r.primaryRole());
                        });
    }

    @FXML
    private void onRemoveMember() {
        if (gridPane == null || shell == null) {
            return;
        }
        String selected = gridPane.selectedMemberName();
        if (selected == null || selected.isBlank()) {
            shell.showWarningDialog("メンバー削除", "削除するメンバー行（氏名列）をクリックして選択してください。");
            return;
        }
        Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
        alert.initOwner(shell.primaryStageForDialogs());
        alert.setTitle("メンバー削除");
        alert.setHeaderText(null);
        alert.setContentText("「" + selected + "」を名簿から削除します。未保存の勤怠セルも行から消えます。");
        if (alert.showAndWait().orElse(ButtonType.CANCEL) != ButtonType.OK) {
            return;
        }
        gridPane.removeMember(selected);
        clearMemberGridCache();
    }

    private boolean confirmSaveWithFourDigit() {
        if (shell == null) {
            return false;
        }
        return FourDigitConfirmationDialog.confirm(
                shell.primaryStageForDialogs(),
                "メンバー勤怠保存",
                "編集内容を attendance-data.json（正本）と 勤怠カレンダー.xlsx に保存します。",
                "保存");
    }

    @FXML
    private void onRefresh() {
        handleUnsavedThen(
                "再読込",
                () -> {
                    clearMemberGridCache();
                    loadGridFromPython();
                    refreshLocalReadiness();
                },
                () -> {});
    }

    @FXML
    private void onOpenAttendanceCalendar() {
        if (shell == null) {
            return;
        }
        Path path = AppPaths.attendanceCalendarXlsxPath(shell.snapshotUiEnv());
        if (!Files.isRegularFile(path)) {
            shell.showWarningDialog(
                    "勤怠・機械カレンダーを開く",
                    "ファイルが見つかりません。\n"
                            + path
                            + "\n「保存」で 勤怠カレンダー.xlsx を出力してから開いてください。");
            return;
        }
        if (!shell.openAttendanceCalendarXlsxInDesktop("[member-attendance]")) {
            shell.showErrorDialog(
                    "勤怠・機械カレンダーを開く",
                    "ファイルを開けませんでした。\n" + path);
        }
    }

    private void loadGridFromPython() {
        loadGridFromPython(null);
    }

    private void loadGridFromPython(Consumer<Boolean> onComplete) {
        if (shell == null) {
            updateGridLoadingOverlay();
            if (onComplete != null) {
                onComplete.accept(false);
            }
            return;
        }
        LocalDate selected = selectedCalendarDate();
        long gen = loadGeneration.incrementAndGet();
        int year = selected.getYear();
        int month = selected.getMonthValue();
        String cacheKey = memberGridCacheKey(year, month);
        JsonNode cached = memberGridCache.get(cacheKey);
        if (cached != null) {
            if (gen == loadGeneration.get()) {
                applyMemberGridNode(cached, year, month);
            }
            if (onComplete != null) {
                onComplete.accept(true);
            }
            return;
        }
        runAsync(
                shell.buildAttendanceDataIoRequest(
                        "member_grid", Integer.toString(year), Integer.toString(month)),
                node -> {
                    if (gen != loadGeneration.get()) {
                        return;
                    }
                    storeMemberGridCache(node);
                    applyMemberGridNode(node, year, month);
                },
                false,
                null,
                gen,
                onComplete,
                year + "年" + month + "月を読込中");
    }

    private void runAsync(
            PythonProcessRunner.RunRequest req,
            java.util.function.Consumer<JsonNode> onOk,
            boolean refreshGridAfter,
            Path tempPatchFile) {
        runAsync(req, onOk, refreshGridAfter, tempPatchFile, null, null, null);
    }

    private void runAsync(
            PythonProcessRunner.RunRequest req,
            java.util.function.Consumer<JsonNode> onOk,
            boolean refreshGridAfter,
            Path tempPatchFile,
            Long gridLoadGen,
            Consumer<Boolean> onFinished) {
        runAsync(req, onOk, refreshGridAfter, tempPatchFile, gridLoadGen, onFinished, null);
    }

    private void runAsync(
            PythonProcessRunner.RunRequest req,
            java.util.function.Consumer<JsonNode> onOk,
            boolean refreshGridAfter,
            Path tempPatchFile,
            Long gridLoadGen,
            Consumer<Boolean> onFinished,
            String loadingMessage) {
        setActiveLoadingMessage(loadingMessage);
        pushTabProcessing();
        PythonProcessRunner.runCaptureAsync(req)
                .whenComplete(
                        (cap, err) ->
                                Platform.runLater(
                                        () -> {
                                            try {
                                                if (tempPatchFile != null) {
                                                    try {
                                                        Files.deleteIfExists(tempPatchFile);
                                                    } catch (Exception ignored) {
                                                        // ignore
                                                    }
                                                }
                                                if (err != null) {
                                                    statusLabel.setText("エラー: " + err.getMessage());
                                                    if (shell != null) {
                                                        shell.appendLog("[member-attendance] " + err);
                                                    }
                                                    if (onFinished != null) {
                                                        onFinished.accept(false);
                                                    }
                                                    return;
                                                }
                                                if (cap == null) {
                                                    statusLabel.setText("失敗");
                                                    if (onFinished != null) {
                                                        onFinished.accept(false);
                                                    }
                                                    return;
                                                }
                                                try {
                                                    JsonNode node =
                                                            JSON.readTree(
                                                                    AttendanceOvertimePreview
                                                                            .MasterReadSummaryJson
                                                                            .extractLastJsonLine(
                                                                                    cap.stdout()));
                                                    if (!node.path("ok").asBoolean(false)) {
                                                        statusLabel.setText(
                                                                "エラー: "
                                                                        + node.path("error")
                                                                                .asText("失敗"));
                                                        if (shell != null) {
                                                            shell.appendLog(
                                                                    "[member-attendance] exit="
                                                                            + cap.exitCode()
                                                                            + " "
                                                                            + cap.stdout());
                                                        }
                                                        if (onFinished != null) {
                                                            onFinished.accept(false);
                                                        }
                                                        return;
                                                    }
                                                    if (cap.exitCode() != 0) {
                                                        statusLabel.setText(
                                                                "失敗 exit=" + cap.exitCode());
                                                        if (shell != null) {
                                                            shell.appendLog(
                                                                    "[member-attendance] "
                                                                            + cap.stdout());
                                                        }
                                                        if (onFinished != null) {
                                                            onFinished.accept(false);
                                                        }
                                                        return;
                                                    }
                                                    if (gridLoadGen != null
                                                            && gridLoadGen
                                                                    != loadGeneration.get()) {
                                                        if (onFinished != null) {
                                                            onFinished.accept(false);
                                                        }
                                                        return;
                                                    }
                                                    onOk.accept(node);
                                                    if (onFinished != null) {
                                                        onFinished.accept(true);
                                                    }
                                                    if (refreshGridAfter) {
                                                        loadGridFromPython();
                                                    }
                                                } catch (Exception e) {
                                                    statusLabel.setText(e.getMessage());
                                                    if (shell != null) {
                                                        shell.appendLog(
                                                                "[member-attendance] " + e);
                                                    }
                                                    if (onFinished != null) {
                                                        onFinished.accept(false);
                                                    }
                                                }
                                            } finally {
                                                popTabProcessing();
                                            }
                                        }));
    }

    private void pushTabProcessing() {
        tabProcessingDepth++;
        if (tabProcessingDepth == 1) {
            setToolbarBusy(true);
            beginStatusActivity();
        }
    }

    private void popTabProcessing() {
        if (tabProcessingDepth > 0) {
            tabProcessingDepth--;
        }
        if (tabProcessingDepth == 0) {
            setToolbarBusy(false);
            endStatusActivity();
            scheduleMemberGridIdlePrefetch();
        }
    }

    private void setActiveLoadingMessage(String message) {
        activeLoadingMessage =
                message != null && !message.isBlank() ? message.strip() : "処理中";
        updateGridLoadingOverlay();
        if (tabProcessingDepth > 0) {
            updateStatusActivityLabel();
        }
    }

    private void beginStatusActivity() {
        statusActivityStartMs = System.currentTimeMillis();
        if (statusProgress != null) {
            statusProgress.setVisible(true);
        }
        updateStatusActivityLabel();
        if (statusActivityTick != null) {
            statusActivityTick.play();
        }
    }

    private void endStatusActivity() {
        if (statusActivityTick != null) {
            statusActivityTick.stop();
        }
        if (statusProgress != null) {
            statusProgress.setVisible(false);
        }
    }

    private void updateStatusActivityLabel() {
        if (statusLabel == null) {
            return;
        }
        double sec = (System.currentTimeMillis() - statusActivityStartMs) / 1000.0;
        statusLabel.setText(
                String.format("%s…（経過 %.1f 秒）", activeLoadingMessage, sec));
    }

    private void finishGridLoadingOverlay(Long gridLoadGen) {
        if (gridLoadGen == null || gridPane == null) {
            return;
        }
        if (gridLoadGen != loadGeneration.get()) {
            return;
        }
        // setToolbarBusy / setGridLoading は popTabProcessing で解除
    }

    private void beginSetupWizardGridOverlay() {
        setupWizardGridOverlayDepth++;
        updateGridLoadingOverlay();
    }

    private void endSetupWizardGridOverlay() {
        if (setupWizardGridOverlayDepth > 0) {
            setupWizardGridOverlayDepth--;
        }
        updateGridLoadingOverlay();
    }

    private void updateGridLoadingOverlay() {
        if (gridPane == null) {
            return;
        }
        boolean loading = setupWizardGridOverlayDepth > 0 || tabProcessingDepth > 0;
        gridPane.setGridLoading(
                loading,
                loading
                        ? (setupWizardGridOverlayDepth > 0
                                ? "セットアップ準備中"
                                : activeLoadingMessage)
                        : null);
    }

    private void setToolbarBusy(boolean busy) {
        if (saveButton != null) {
            saveButton.setDisable(busy);
        }
        if (setupButton != null) {
            setupButton.setDisable(busy);
        }
        if (restoreButton != null) {
            restoreButton.setDisable(busy);
        }
        if (refreshButton != null) {
            refreshButton.setDisable(busy);
        }
        if (openCalendarButton != null) {
            openCalendarButton.setDisable(busy);
        }
        if (cellSizeSpinner != null) {
            cellSizeSpinner.setDisable(busy);
        }
        if (monthCalendar != null) {
            monthCalendar.setNavigationEnabled(!busy);
        }
        updateGridLoadingOverlay();
    }
}
