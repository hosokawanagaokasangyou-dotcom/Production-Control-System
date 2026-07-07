package jp.co.pm.ai.desktop;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;

import javafx.application.Platform;
import javafx.beans.binding.Bindings;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.collections.transformation.FilteredList;
import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonBar;
import javafx.scene.control.ButtonType;
import javafx.scene.control.CheckBox;
import javafx.scene.control.Label;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.input.Clipboard;
import javafx.scene.input.ClipboardContent;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup;
import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup.PlanEntry;
import jp.co.pm.ai.desktop.ui.ClipboardTableSupport;
import jp.co.pm.ai.desktop.ui.ColumnVisibilitySupport;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;
import jp.co.pm.ai.desktop.ui.TableViewColumnSettingsStrip;
import jp.co.pm.ai.desktop.reconciliation.JuchuTransferCoverageCheck.ColumnCheck;
import jp.co.pm.ai.desktop.reconciliation.JuchuHeaderAliasRegistry;
import jp.co.pm.ai.desktop.reconciliation.KonanDailyReportLookup;
import jp.co.pm.ai.desktop.reconciliation.RemoteDesktopLatestSourceFiles;
import jp.co.pm.ai.desktop.reconciliation.RequestFormPipelineStatusService;
import jp.co.pm.ai.desktop.reconciliation.RequestFormPipelineStatusService.PipelineStatusRow;
import jp.co.pm.ai.desktop.reconciliation.RequestFormPipelineStatusService.ScanResult;

/** 依頼書原本の受注転記率・アラジン加工計画の有無を一覧表示する。 */
public final class RequestFormPipelineCheckTabController {

    private static final int PLAN_DAY_COLUMNS =
            AladdinShapedPlanQtyLookup.PIPELINE_CHECK_PLAN_DAY_COLUMNS;

    private static final String HINT_TEXT =
            "依頼書原本フォルダ内の Excel 原本と TPI PDF を走査し、受注ファイルへの転記状況と"
                    + " shaped_aladdin_plan.json 上のアラジン加工計画を照合します。"
                    + " 転記率は原本に値がある転記対象列を分母とします。"
                    + " ①〜⑦列は依頼ごとの加工計画日（昇順・最大7日）を表示（セル例: 7/3 100m）。"
                    + " JSON表の日付列順とは無関係。"
                    + " 受注入力日フィルタは既定30日（変更可）。"
                    + " 「日報」列は加工日報の完了区分を依頼単位で集約（完了/未了/―）。"
                    + " 「日報未了のみ表示」は既定オン（完了済み依頼を非表示）。"
                    + " 下段表はアラジン計画が無い場合、加工日報の工程行を表示。";

    private record MainColDef(String title, String property, double defaultWidth) {}

    private static List<MainColDef> defaultMainColumnDefs() {
        List<MainColDef> defs = new ArrayList<>();
        defs.add(new MainColDef("依頼No", "iraiNo", 90));
        defs.add(new MainColDef("原本", "originalFile", 140));
        defs.add(new MainColDef("回答日", "indexResponseDate", 88));
        defs.add(new MainColDef("原反投入日", "indexInputDate", 88));
        defs.add(new MainColDef("納期", "indexDeliveryDate", 88));
        defs.add(new MainColDef("納期備考", "indexDeliveryRemarks", 120));
        defs.add(new MainColDef("目次契約NO", "indexContractNo", 120));
        defs.add(new MainColDef("契約備考", "indexContractRemarks", 120));
        defs.add(new MainColDef("入力日", "juchuInputDate", 88));
        defs.add(new MainColDef("入力者", "juchuInputOperator", 88));
        defs.add(new MainColDef("調整納期", "juchuAdjustDeliveryDate", 88));
        defs.add(new MainColDef("転記率", "rateDisplay", 100));
        defs.add(new MainColDef("未一致", "mismatchCount", 52));
        defs.add(new MainColDef("原本契約NO", "originalContractNo", 120));
        defs.add(new MainColDef("契約NO", "contractNoStatus", 120));
        defs.add(new MainColDef("Aladdin", "aladdinStatus", 64));
        defs.add(new MainColDef("日報", "dailyReportOrderStatus", 52));
        for (int i = 0; i < PLAN_DAY_COLUMNS; i++) {
            defs.add(
                    new MainColDef(
                            AladdinShapedPlanQtyLookup.circledSlotColumnLabel(i),
                            "planDay" + i,
                            88));
        }
        return List.copyOf(defs);
    }

    public static final class MainRow {
        private String iraiNo;
        private String originalFile;
        private String indexResponseDate;
        private String indexInputDate;
        private String indexDeliveryDate;
        private String indexDeliveryRemarks;
        private String indexContractNo;
        private String indexContractRemarks;
        private String juchuInputDate;
        private String juchuInputOperator;
        private String juchuAdjustDeliveryDate;
        private String rateDisplay;
        private String mismatchCount;
        private String originalContractNo;
        private String contractNoStatus;
        private String aladdinStatus;
        /** 加工日報の依頼単位ステータス: 完了 / 未了 / ― */
        private String dailyReportOrderStatus;
        private final List<String> planDayValues = new ArrayList<>(PLAN_DAY_COLUMNS);
        private PipelineStatusRow source;

        public MainRow() {
            for (int i = 0; i < PLAN_DAY_COLUMNS; i++) {
                planDayValues.add("");
            }
        }

        public String getIraiNo() {
            return iraiNo;
        }

        public void setIraiNo(String iraiNo) {
            this.iraiNo = iraiNo;
        }

        public String getOriginalFile() {
            return originalFile;
        }

        public void setOriginalFile(String originalFile) {
            this.originalFile = originalFile;
        }

        public String getIndexResponseDate() {
            return indexResponseDate;
        }

        public void setIndexResponseDate(String indexResponseDate) {
            this.indexResponseDate = indexResponseDate;
        }

        public String getIndexInputDate() {
            return indexInputDate;
        }

        public void setIndexInputDate(String indexInputDate) {
            this.indexInputDate = indexInputDate;
        }

        public String getIndexDeliveryDate() {
            return indexDeliveryDate;
        }

        public void setIndexDeliveryDate(String indexDeliveryDate) {
            this.indexDeliveryDate = indexDeliveryDate;
        }

        public String getIndexDeliveryRemarks() {
            return indexDeliveryRemarks;
        }

        public void setIndexDeliveryRemarks(String indexDeliveryRemarks) {
            this.indexDeliveryRemarks = indexDeliveryRemarks;
        }

        public String getIndexContractNo() {
            return indexContractNo;
        }

        public void setIndexContractNo(String indexContractNo) {
            this.indexContractNo = indexContractNo;
        }

        public String getIndexContractRemarks() {
            return indexContractRemarks;
        }

        public void setIndexContractRemarks(String indexContractRemarks) {
            this.indexContractRemarks = indexContractRemarks;
        }

        public String getJuchuInputDate() {
            return juchuInputDate;
        }

        public void setJuchuInputDate(String juchuInputDate) {
            this.juchuInputDate = juchuInputDate;
        }

        public String getJuchuInputOperator() {
            return juchuInputOperator;
        }

        public void setJuchuInputOperator(String juchuInputOperator) {
            this.juchuInputOperator = juchuInputOperator;
        }

        public String getJuchuAdjustDeliveryDate() {
            return juchuAdjustDeliveryDate;
        }

        public void setJuchuAdjustDeliveryDate(String juchuAdjustDeliveryDate) {
            this.juchuAdjustDeliveryDate = juchuAdjustDeliveryDate;
        }

        public String getRateDisplay() {
            return rateDisplay;
        }

        public void setRateDisplay(String rateDisplay) {
            this.rateDisplay = rateDisplay;
        }

        public String getMismatchCount() {
            return mismatchCount;
        }

        public void setMismatchCount(String mismatchCount) {
            this.mismatchCount = mismatchCount;
        }

        public String getOriginalContractNo() {
            return originalContractNo;
        }

        public void setOriginalContractNo(String originalContractNo) {
            this.originalContractNo = originalContractNo;
        }

        public String getContractNoStatus() {
            return contractNoStatus;
        }

        public void setContractNoStatus(String contractNoStatus) {
            this.contractNoStatus = contractNoStatus;
        }

        public String getAladdinStatus() {
            return aladdinStatus;
        }

        public void setAladdinStatus(String aladdinStatus) {
            this.aladdinStatus = aladdinStatus;
        }

        public String getDailyReportOrderStatus() {
            return dailyReportOrderStatus;
        }

        public void setDailyReportOrderStatus(String dailyReportOrderStatus) {
            this.dailyReportOrderStatus = dailyReportOrderStatus;
        }

        public String getPlanDay0() {
            return planDayAt(0);
        }

        public String getPlanDay1() {
            return planDayAt(1);
        }

        public String getPlanDay2() {
            return planDayAt(2);
        }

        public String getPlanDay3() {
            return planDayAt(3);
        }

        public String getPlanDay4() {
            return planDayAt(4);
        }

        public String getPlanDay5() {
            return planDayAt(5);
        }

        public String getPlanDay6() {
            return planDayAt(6);
        }

        public void setPlanDayValues(List<String> values) {
            for (int i = 0; i < PLAN_DAY_COLUMNS; i++) {
                planDayValues.set(i, values != null && i < values.size() ? nullToEmpty(values.get(i)) : "");
            }
        }

        public PipelineStatusRow source() {
            return source;
        }

        public void setSource(PipelineStatusRow source) {
            this.source = source;
        }

        private String planDayAt(int index) {
            return index >= 0 && index < planDayValues.size() ? planDayValues.get(index) : "";
        }

        private static String nullToEmpty(String val) {
            return val != null ? val : "";
        }
    }

    public static final class MismatchRow {
        private final SimpleStringProperty formLabel = new SimpleStringProperty();
        private final SimpleStringProperty originalValue = new SimpleStringProperty();
        private final SimpleStringProperty juchuValue = new SimpleStringProperty();

        MismatchRow(ColumnCheck check) {
            formLabel.set(check.formLabel());
            originalValue.set(check.originalValue());
            juchuValue.set(check.juchuValue());
        }

        public String getFormLabel() {
            return formLabel.get();
        }

        public SimpleStringProperty formLabelProperty() {
            return formLabel;
        }

        public String getOriginalValue() {
            return originalValue.get();
        }

        public SimpleStringProperty originalValueProperty() {
            return originalValue;
        }

        public String getJuchuValue() {
            return juchuValue.get();
        }

        public SimpleStringProperty juchuValueProperty() {
            return juchuValue;
        }
    }

    public static final class PlanRow {
        private final SimpleStringProperty machineName = new SimpleStringProperty();
        private final SimpleStringProperty processName = new SimpleStringProperty();
        private final SimpleStringProperty dateYmd = new SimpleStringProperty();
        private final SimpleStringProperty planMeters = new SimpleStringProperty();
        private final SimpleStringProperty dailyReportCompletion = new SimpleStringProperty();

        PlanRow(PlanEntry entry, String dailyReportCompletionDisplay) {
            machineName.set(entry.machineName());
            processName.set(entry.processName());
            dateYmd.set(entry.dateYmd());
            planMeters.set(formatMeters(entry.planMeters()));
            dailyReportCompletion.set(
                    dailyReportCompletionDisplay != null ? dailyReportCompletionDisplay : "");
        }

        public String getMachineName() {
            return machineName.get();
        }

        public SimpleStringProperty machineNameProperty() {
            return machineName;
        }

        public String getProcessName() {
            return processName.get();
        }

        public SimpleStringProperty processNameProperty() {
            return processName;
        }

        public String getDateYmd() {
            return dateYmd.get();
        }

        public SimpleStringProperty dateYmdProperty() {
            return dateYmd;
        }

        public String getPlanMeters() {
            return planMeters.get();
        }

        public SimpleStringProperty planMetersProperty() {
            return planMeters;
        }

        public String getDailyReportCompletion() {
            return dailyReportCompletion.get();
        }

        public SimpleStringProperty dailyReportCompletionProperty() {
            return dailyReportCompletion;
        }

        private static String formatMeters(double m) {
            if (Math.abs(m - Math.rint(m)) < 1e-9) {
                return String.valueOf((long) Math.rint(m));
            }
            return String.valueOf(m);
        }
    }

    /** クリップボード用の依頼ヘッダ（上段一覧の選択行）。 */
    record PlanCopyHeader(String iraiNo, String contractNo, String rawInputDate) {

        static PlanCopyHeader from(MainRow row) {
            if (row == null) {
                return null;
            }
            String rawInputDate = "";
            PipelineStatusRow src = row.source();
            if (src != null && src.rawInputDateDisplay() != null) {
                rawInputDate = src.rawInputDateDisplay();
            }
            return new PlanCopyHeader(
                    nullToEmpty(row.getIraiNo()),
                    nullToEmpty(row.getContractNoStatus()),
                    nullToEmpty(rawInputDate));
        }

        private static String nullToEmpty(String val) {
            return val != null ? val : "";
        }
    }

    @FXML
    private Button refreshButton;

    @FXML
    private ProgressIndicator refreshProgressIndicator;

    @FXML
    private TextField filterField;

    @FXML
    private CheckBox hideNoOriginalCheck;

    @FXML
    private CheckBox hideOldJuchuInputCheck;

    @FXML
    private Spinner<Integer> juchuInputHideDaysSpinner;

    @FXML
    private CheckBox showNoAladdinCheck;

    @FXML
    private CheckBox showAdjustDeliveryFromTodayCheck;

    @FXML
    private CheckBox showDailyReportIncompleteOnlyCheck;

    @FXML
    private Label statusLabel;

    @FXML
    private Label hintLabel;

    @FXML
    private Label dailyReportSourceLabel;

    @FXML
    private TableView<MainRow> mainTable;

    @FXML
    private Button copyMainTableButton;

    @FXML
    private Button copyMainTableForEmailButton;

    @FXML
    private Button copyMainTableHtmlButton;

    @FXML
    private Button copyMainTableCsvButton;

    @FXML
    private HBox mainColumnStripHost;

    @FXML
    private TableView<MismatchRow> mismatchTable;

    @FXML
    private TableView<PlanRow> planTable;

    @FXML
    private Button copyPlanTableButton;

    @FXML
    private Button copyPlanTableForEmailButton;

    private MainShellController shell;

    private final ObservableList<MainRow> allRows = FXCollections.observableArrayList();
    private FilteredList<MainRow> filteredRows;
    private final ObservableList<MismatchRow> mismatchRows = FXCollections.observableArrayList();
    private final ObservableList<PlanRow> planRows = FXCollections.observableArrayList();

    /** 下段アラジン計画表の表示元（コピー時の依頼ヘッダ）。 */
    private MainRow planContextRow;

    private KonanDailyReportLookup dailyReportLookup = KonanDailyReportLookup.empty();

    private boolean aladdinJsonAvailable = true;
    private String lastScanWarnings = "";
    /** 起動後に一度でも走査結果を反映したら true（手動「更新」は常に可）。 */
    private boolean scanApplied;
    private boolean refreshInProgress;

    private final List<MainColDef> mainColumnDefs = new ArrayList<>(defaultMainColumnDefs());

    private final Map<String, String> mainColumnTitleToProperty = new LinkedHashMap<>();

    private final AtomicBoolean suppressColumnPersistence = new AtomicBoolean(false);

    private boolean mainColumnsInitialized;

    @FXML
    private void initialize() {
        hintLabel.setText(HINT_TEXT);
        updateDailyReportSourceLabel();
        filteredRows = new FilteredList<>(allRows, row -> true);
        mainTable.setItems(filteredRows);
        mismatchTable.setItems(mismatchRows);
        mismatchTable.setPlaceholder(new Label("相違なし"));
        planTable.setItems(planRows);
        VBox.setVgrow(mainTable, Priority.ALWAYS);

        setupMismatchColumns();
        setupPlanColumns();

        if (copyMainTableButton != null) {
            copyMainTableButton.disableProperty().bind(Bindings.isEmpty(filteredRows));
        }
        if (copyMainTableForEmailButton != null) {
            copyMainTableForEmailButton.disableProperty().bind(Bindings.isEmpty(filteredRows));
        }
        if (copyMainTableHtmlButton != null) {
            copyMainTableHtmlButton.disableProperty().bind(Bindings.isEmpty(filteredRows));
        }
        if (copyMainTableCsvButton != null) {
            copyMainTableCsvButton.disableProperty().bind(Bindings.isEmpty(filteredRows));
        }
        if (copyPlanTableButton != null) {
            copyPlanTableButton.disableProperty().bind(Bindings.isEmpty(planRows));
        }
        if (copyPlanTableForEmailButton != null) {
            copyPlanTableForEmailButton.disableProperty().bind(Bindings.isEmpty(planRows));
        }

        filterField.textProperty().addListener((obs, oldVal, newVal) -> applyFilter());
        hideNoOriginalCheck.selectedProperty().addListener((obs, o, n) -> applyFilter());
        if (showNoAladdinCheck != null) {
            showNoAladdinCheck.selectedProperty().addListener((obs, o, n) -> applyFilter());
        }
        if (showAdjustDeliveryFromTodayCheck != null) {
            showAdjustDeliveryFromTodayCheck
                    .selectedProperty()
                    .addListener((obs, o, n) -> applyFilter());
        }
        if (showDailyReportIncompleteOnlyCheck != null) {
            showDailyReportIncompleteOnlyCheck
                    .selectedProperty()
                    .addListener((obs, o, n) -> applyFilter());
        }
        if (juchuInputHideDaysSpinner != null) {
            juchuInputHideDaysSpinner.setValueFactory(
                    new SpinnerValueFactory.IntegerSpinnerValueFactory(
                            0,
                            999,
                            RequestFormPipelineStatusService.DEFAULT_JUCHU_INPUT_DATE_HIDE_DAYS));
            juchuInputHideDaysSpinner.setEditable(true);
            juchuInputHideDaysSpinner
                    .valueProperty()
                    .addListener((obs, o, n) -> applyFilter());
        }
        if (hideOldJuchuInputCheck != null) {
            hideOldJuchuInputCheck
                    .selectedProperty()
                    .addListener(
                            (obs, o, n) -> {
                                if (juchuInputHideDaysSpinner != null) {
                                    juchuInputHideDaysSpinner.setDisable(!n);
                                }
                                applyFilter();
                            });
            if (juchuInputHideDaysSpinner != null) {
                juchuInputHideDaysSpinner.setDisable(!hideOldJuchuInputCheck.isSelected());
            }
        }

        mainTable.getSelectionModel()
                .selectedItemProperty()
                .addListener((obs, oldRow, newRow) -> showDetail(newRow));
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        setupMainColumnsOnce();
    }

    /** メインシェルで当該タブが選択されたとき。起動後未走査なら自動更新する。 */
    void onMainShellTabSelected() {
        scheduleInitialRefreshIfNeeded();
    }

    private void scheduleInitialRefreshIfNeeded() {
        if (scanApplied || refreshInProgress || shell == null) {
            return;
        }
        Platform.runLater(this::startRefresh);
    }

    @FXML
    private void onRefreshButtonAction() {
        if (!confirmRefreshDespiteStaleSources()) {
            return;
        }
        startRefresh();
    }

    /**
     * 加工計画・加工日報の取得元が {@link RemoteDesktopLatestSourceFiles#PIPELINE_CHECK_MAX_SOURCE_AGE_MS}
     * より古いとき警告し、続行可否を返す。
     */
    private boolean confirmRefreshDespiteStaleSources() {
        if (shell == null) {
            return true;
        }
        List<RemoteDesktopLatestSourceFiles.Row> stale =
                RemoteDesktopLatestSourceFiles.pipelineCheckSourcesOlderThan(
                        shell.snapshotUiEnv(),
                        RemoteDesktopLatestSourceFiles.PIPELINE_CHECK_MAX_SOURCE_AGE_MS,
                        System.currentTimeMillis());
        if (stale.isEmpty()) {
            return true;
        }
        StringBuilder body = new StringBuilder();
        body.append("加工計画または加工日報の取得元が1時間以上更新されていません。");
        body.append(" リモートデスクトップで RPA 取得後に更新することを推奨します。");
        body.append("\n\n");
        for (RemoteDesktopLatestSourceFiles.Row row : stale) {
            body.append("■ ").append(row.category().label()).append('\n');
            body.append("  取得日時: ").append(row.displayAcquiredAt()).append('\n');
            body.append("  ").append(row.displayPath()).append("\n\n");
        }
        body.append("このまま走査を続行しますか？");

        ButtonType proceed = new ButtonType("更新を続行", ButtonBar.ButtonData.OK_DONE);
        ButtonType cancel = new ButtonType("中止", ButtonBar.ButtonData.CANCEL_CLOSE);
        Alert alert = new Alert(Alert.AlertType.WARNING, "", proceed, cancel);
        alert.setTitle("取得データが古い可能性");
        alert.setHeaderText("加工計画／加工日報の取得元が1時間以上前です");
        Label message = new Label(body.toString());
        message.setWrapText(true);
        message.setMaxWidth(560);
        message.setMinWidth(400);
        alert.getDialogPane().setContent(message);
        return alert.showAndWait().orElse(cancel) == proceed;
    }

    private void startRefresh() {
        if (shell == null) {
            statusLabel.setText("シェル未接続");
            return;
        }
        if (refreshInProgress) {
            return;
        }
        refreshInProgress = true;
        setRefreshing(true);
        statusLabel.setText("走査中…");
        JuchuHeaderAliasRegistry registry = shell.snapshotJuchuHeaderAliasRegistryForExport();
        Thread worker =
                new Thread(
                        () -> {
                            try {
                                ScanResult result =
                                        RequestFormPipelineStatusService.scan(
                                                shell.snapshotUiEnv(), registry);
                                Platform.runLater(
                                        () -> {
                                            setRefreshing(false);
                                            applyScanResult(result);
                                        });
                            } catch (Exception ex) {
                                Platform.runLater(
                                        () -> {
                                            setRefreshing(false);
                                            statusLabel.setText(
                                                    "走査失敗: "
                                                            + (ex.getMessage() != null
                                                                    ? ex.getMessage()
                                                                    : ex.toString()));
                                        });
                            }
                        },
                        "request-form-pipeline-check");
        worker.setDaemon(true);
        worker.start();
    }

    private void setRefreshing(boolean on) {
        if (!on) {
            refreshInProgress = false;
        }
        if (refreshProgressIndicator != null) {
            refreshProgressIndicator.setVisible(on);
            refreshProgressIndicator.setManaged(on);
        }
        if (refreshButton != null) {
            refreshButton.setDisable(on);
        }
    }

    private void applyScanResult(ScanResult result) {
        scanApplied = true;
        allRows.clear();
        aladdinJsonAvailable = result.aladdinJsonAvailable();
        dailyReportLookup =
                result.dailyReportLookup() != null
                        ? result.dailyReportLookup()
                        : KonanDailyReportLookup.empty();

        for (PipelineStatusRow row : result.rows()) {
            MainRow ui = new MainRow();
            ui.setIraiNo(row.iraiNo());
            ui.setOriginalFile(
                    row.originalPresent()
                            ? row.originalFileName()
                            : "（依頼書原本なし）");
            ui.setIndexResponseDate(nullToEmpty(row.indexResponseDate()));
            ui.setIndexInputDate(nullToEmpty(row.indexInputDate()));
            ui.setIndexDeliveryDate(nullToEmpty(row.indexDeliveryDate()));
            ui.setIndexDeliveryRemarks(nullToEmpty(row.indexDeliveryRemarks()));
            ui.setIndexContractNo(nullToEmpty(row.indexContractNo()));
            ui.setIndexContractRemarks(nullToEmpty(row.indexContractRemarks()));
            ui.setJuchuInputDate(
                    row.juchuInputDateDisplay() != null ? row.juchuInputDateDisplay() : "");
            ui.setJuchuInputOperator(
                    row.juchuInputOperatorDisplay() != null ? row.juchuInputOperatorDisplay() : "");
            ui.setJuchuAdjustDeliveryDate(
                    row.juchuAdjustDeliveryDateDisplay() != null
                            ? row.juchuAdjustDeliveryDateDisplay()
                            : "");
            ui.setRateDisplay(row.rateDisplay());
            ui.setMismatchCount(String.valueOf(row.mismatchCount()));
            ui.setOriginalContractNo(
                    row.originalContractNoDisplay() != null
                            ? row.originalContractNoDisplay()
                            : "");
            ui.setContractNoStatus(row.contractNoStatus());
            if (!aladdinJsonAvailable) {
                ui.setAladdinStatus("未確認");
            } else {
                ui.setAladdinStatus(row.aladdinPresent() ? "あり" : "なし");
            }
            ui.setDailyReportOrderStatus(
                    dailyReportLookup.orderCompletionStatus(row.iraiNo()));
            ui.setPlanDayValues(row.planDayValues());
            ui.setSource(row);
            allRows.add(ui);
        }
        lastScanWarnings =
                result.warnings().isEmpty() ? "" : String.join(" | ", result.warnings());
        applyFilter();
        updateStatusLabel();
        updateDailyReportSourceLabel();

        if (shell != null) {
            shell.appendLog(
                    "[pipeline-check] "
                            + result.rows().size()
                            + " 件"
                            + (result.warnings().isEmpty()
                                    ? ""
                                    : " warnings=" + result.warnings().size()));
        }
        mainTable.getSelectionModel().clearSelection();
        mismatchRows.clear();
        planRows.clear();
        planContextRow = null;
    }

    private void showDetail(MainRow row) {
        mismatchRows.clear();
        planRows.clear();
        planContextRow = row;
        if (row == null || row.source() == null) {
            return;
        }
        PipelineStatusRow src = row.source();
        if (src.coverage() != null) {
            for (ColumnCheck check : src.coverage().details()) {
                if (!check.matched()) {
                    mismatchRows.add(new MismatchRow(check));
                }
            }
        }
        for (PlanEntry entry : src.planEntries()) {
            String completion =
                    dailyReportLookup.completionDisplay(
                            src.iraiNo(), entry.processName(), entry.machineName());
            planRows.add(new PlanRow(entry, completion));
        }
        if (planRows.isEmpty()) {
            for (KonanDailyReportLookup.OrderDailyReportEntry dr :
                    dailyReportLookup.entriesForOrder(src.iraiNo())) {
                planRows.add(
                        new PlanRow(
                                new PlanEntry(
                                        dr.machineName(),
                                        dr.processName(),
                                        "",
                                        0.0),
                                dr.completionStatus()));
            }
        }
    }

    @FXML
    private void onCopyMainTableButtonAction() {
        if (filteredRows.isEmpty()) {
            return;
        }
        List<MainRow> rows = List.copyOf(filteredRows);
        List<String> titles = mainTableVisibleColumnTitles();
        String tsv = formatMainTableTsv(rows, titles);
        ClipboardContent content = new ClipboardContent();
        content.putString(tsv);
        Clipboard.getSystemClipboard().setContent(content);
        if (shell != null) {
            shell.appendLog(
                    "[pipeline-check] 一覧 "
                            + rows.size()
                            + " 行をクリップボードへ表としてコピー");
        }
    }

    @FXML
    private void onCopyMainTableForEmailButtonAction() {
        if (filteredRows.isEmpty()) {
            return;
        }
        List<MainRow> rows = List.copyOf(filteredRows);
        List<String> titles = mainTableVisibleColumnTitles();
        String tsv = formatMainTableTsv(rows, titles);
        String html = formatMainTableHtml(rows, titles);
        ClipboardTableSupport.copyTabularForRichTextPaste(tsv, html);
        if (shell != null) {
            shell.appendLog(
                    "[pipeline-check] 一覧 "
                            + rows.size()
                            + " 行をメール貼り付け用（HTML表）でクリップボードへコピー");
        }
    }

    @FXML
    private void onCopyMainTableHtmlButtonAction() {
        if (filteredRows.isEmpty()) {
            return;
        }
        List<MainRow> rows = List.copyOf(filteredRows);
        List<String> titles = mainTableVisibleColumnTitles();
        String html = formatMainTableHtml(rows, titles);
        ClipboardTableSupport.copyHtmlTableOnly(html);
        if (shell != null) {
            shell.appendLog(
                    "[pipeline-check] 一覧 "
                            + rows.size()
                            + " 行を HTML形式でクリップボードへコピー");
        }
    }

    @FXML
    private void onCopyMainTableCsvButtonAction() {
        if (filteredRows.isEmpty()) {
            return;
        }
        List<MainRow> rows = List.copyOf(filteredRows);
        List<String> titles = mainTableVisibleColumnTitles();
        String csv = formatMainTableCsv(rows, titles);
        ClipboardContent content = new ClipboardContent();
        content.putString(csv);
        Clipboard.getSystemClipboard().setContent(content);
        if (shell != null) {
            shell.appendLog(
                    "[pipeline-check] 一覧 "
                            + rows.size()
                            + " 行をカンマ区切りでクリップボードへコピー");
        }
    }

    @FXML
    private void onCopyPlanTableButtonAction() {
        if (planRows.isEmpty()) {
            return;
        }
        PlanCopyHeader header = PlanCopyHeader.from(planContextRow);
        String tsv = formatPlanTableTsv(header, planRows);
        ClipboardContent content = new ClipboardContent();
        content.putString(tsv);
        Clipboard.getSystemClipboard().setContent(content);
        if (shell != null) {
            shell.appendLog(
                    "[pipeline-check] アラジン加工計画 "
                            + planRows.size()
                            + " 行をクリップボードへ表としてコピー"
                            + headerLogSuffix(header));
        }
    }

    @FXML
    private void onCopyPlanTableForEmailButtonAction() {
        if (planRows.isEmpty()) {
            return;
        }
        PlanCopyHeader header = PlanCopyHeader.from(planContextRow);
        String tsv = formatPlanTableTsv(header, planRows);
        String html = formatPlanTableHtml(header, planRows);
        ClipboardTableSupport.copyTabularForRichTextPaste(tsv, html);
        if (shell != null) {
            shell.appendLog(
                    "[pipeline-check] アラジン加工計画 "
                            + planRows.size()
                            + " 行をメール貼り付け用（HTML表）でクリップボードへコピー"
                            + headerLogSuffix(header));
        }
    }

    private static String headerLogSuffix(PlanCopyHeader header) {
        if (header == null || header.iraiNo().isBlank()) {
            return "";
        }
        return "（依頼No: " + header.iraiNo() + "）";
    }

    static String formatPlanTableHtml(PlanCopyHeader header, List<PlanRow> rows) {
        StringBuilder sb = new StringBuilder();
        appendPlanCopyHeaderHtml(sb, header);
        sb.append(
                "<table border=\"1\" cellspacing=\"0\" cellpadding=\"4\""
                        + " style=\"border-collapse:collapse;font-family:'Meiryo UI',sans-serif;font-size:11pt;\">");
        sb.append("<thead><tr>");
        for (String columnTitle : List.of("機械名", "工程名", "日付", "計画m", "日報完了")) {
            sb.append("<th style=\"background:#D9E1F2;padding:4px 8px;text-align:left;\">")
                    .append(ClipboardTableSupport.escapeHtml(columnTitle))
                    .append("</th>");
        }
        sb.append("</tr></thead><tbody>");
        for (PlanRow row : rows) {
            sb.append("<tr>");
            appendHtmlCell(sb, row.getMachineName());
            appendHtmlCell(sb, row.getProcessName());
            appendHtmlCell(sb, row.getDateYmd());
            appendHtmlCell(sb, row.getPlanMeters());
            appendHtmlCell(sb, row.getDailyReportCompletion());
            sb.append("</tr>");
        }
        sb.append("</tbody></table>");
        return sb.toString();
    }

    private static void appendHtmlCell(StringBuilder sb, String value) {
        sb.append("<td style=\"padding:4px 8px;\">")
                .append(ClipboardTableSupport.escapeHtml(value != null ? value : ""))
                .append("</td>");
    }

    static String formatMainTableTsv(List<MainRow> rows) {
        return formatMainTableTsv(rows, mainTableColumnTitles());
    }

    static String formatMainTableTsv(List<MainRow> rows, List<String> titles) {
        List<String> headers = titles != null && !titles.isEmpty() ? titles : mainTableColumnTitles();
        StringBuilder sb = new StringBuilder();
        appendMainTableTsvRow(sb, headers);
        for (MainRow row : rows) {
            sb.append('\n');
            appendMainTableTsvRow(sb, mainTableRowCells(row, headers));
        }
        return sb.toString();
    }

    static String formatMainTableHtml(List<MainRow> rows) {
        return formatMainTableHtml(rows, mainTableColumnTitles());
    }

    static String formatMainTableHtml(List<MainRow> rows, List<String> titles) {
        List<String> headers = titles != null && !titles.isEmpty() ? titles : mainTableColumnTitles();
        StringBuilder sb = new StringBuilder();
        sb.append(
                "<table border=\"1\" cellspacing=\"0\" cellpadding=\"4\""
                        + " style=\"border-collapse:collapse;font-family:'Meiryo UI',sans-serif;font-size:11pt;\">");
        sb.append("<thead><tr>");
        for (String columnTitle : headers) {
            sb.append("<th style=\"background:#D9E1F2;padding:4px 8px;text-align:left;\">")
                    .append(ClipboardTableSupport.escapeHtml(columnTitle))
                    .append("</th>");
        }
        sb.append("</tr></thead><tbody>");
        for (MainRow row : rows) {
            sb.append("<tr>");
            for (String cell : mainTableRowCells(row, headers)) {
                appendHtmlCell(sb, cell);
            }
            sb.append("</tr>");
        }
        sb.append("</tbody></table>");
        return sb.toString();
    }

    static String formatMainTableCsv(List<MainRow> rows) {
        return formatMainTableCsv(rows, mainTableColumnTitles());
    }

    static String formatMainTableCsv(List<MainRow> rows, List<String> titles) {
        List<String> headers = titles != null && !titles.isEmpty() ? titles : mainTableColumnTitles();
        StringBuilder sb = new StringBuilder();
        appendMainTableCsvRow(sb, headers);
        for (MainRow row : rows) {
            sb.append('\n');
            appendMainTableCsvRow(sb, mainTableRowCells(row, headers));
        }
        return sb.toString();
    }

    static List<String> mainTableColumnTitles() {
        List<String> titles = new ArrayList<>();
        for (MainColDef def : defaultMainColumnDefs()) {
            titles.add(def.title());
        }
        return List.copyOf(titles);
    }

    static List<String> mainTableRowCells(MainRow row) {
        return mainTableRowCells(row, mainTableColumnTitles());
    }

    static List<String> mainTableRowCells(MainRow row, List<String> titles) {
        if (row == null) {
            return List.of();
        }
        Map<String, String> titleToProperty = defaultMainColumnTitleToProperty();
        List<String> cells = new ArrayList<>();
        for (String title : titles) {
            String property = titleToProperty.get(title);
            cells.add(property != null ? mainRowCellByProperty(row, property) : "");
        }
        return List.copyOf(cells);
    }

    private static Map<String, String> defaultMainColumnTitleToProperty() {
        Map<String, String> map = new LinkedHashMap<>();
        for (MainColDef def : defaultMainColumnDefs()) {
            map.put(def.title(), def.property());
        }
        return map;
    }

    static String mainRowCellByProperty(MainRow row, String property) {
        if (row == null || property == null) {
            return "";
        }
        return nullToEmpty(
                switch (property) {
                    case "iraiNo" -> row.getIraiNo();
                    case "originalFile" -> row.getOriginalFile();
                    case "indexResponseDate" -> row.getIndexResponseDate();
                    case "indexInputDate" -> row.getIndexInputDate();
                    case "indexDeliveryDate" -> row.getIndexDeliveryDate();
                    case "indexDeliveryRemarks" -> row.getIndexDeliveryRemarks();
                    case "indexContractNo" -> row.getIndexContractNo();
                    case "indexContractRemarks" -> row.getIndexContractRemarks();
                    case "juchuInputDate" -> row.getJuchuInputDate();
                    case "juchuInputOperator" -> row.getJuchuInputOperator();
                    case "juchuAdjustDeliveryDate" -> row.getJuchuAdjustDeliveryDate();
                    case "rateDisplay" -> row.getRateDisplay();
                    case "mismatchCount" -> row.getMismatchCount();
                    case "originalContractNo" -> row.getOriginalContractNo();
                    case "contractNoStatus" -> row.getContractNoStatus();
                    case "aladdinStatus" -> row.getAladdinStatus();
                    case "dailyReportOrderStatus" -> row.getDailyReportOrderStatus();
                    case "planDay0" -> row.getPlanDay0();
                    case "planDay1" -> row.getPlanDay1();
                    case "planDay2" -> row.getPlanDay2();
                    case "planDay3" -> row.getPlanDay3();
                    case "planDay4" -> row.getPlanDay4();
                    case "planDay5" -> row.getPlanDay5();
                    case "planDay6" -> row.getPlanDay6();
                    default -> "";
                });
    }

    private static String nullToEmpty(String value) {
        return value != null ? value : "";
    }

    private static void appendMainTableTsvRow(StringBuilder sb, List<String> cells) {
        for (int i = 0; i < cells.size(); i++) {
            if (i > 0) {
                sb.append('\t');
            }
            appendTsvCell(sb, cells.get(i));
        }
    }

    private static void appendMainTableCsvRow(StringBuilder sb, List<String> cells) {
        for (int i = 0; i < cells.size(); i++) {
            if (i > 0) {
                sb.append(',');
            }
            appendCsvCell(sb, cells.get(i));
        }
    }

    static String formatPlanTableTsv(PlanCopyHeader header, List<PlanRow> rows) {
        StringBuilder sb = new StringBuilder();
        appendPlanCopyHeaderTsv(sb, header);
        if (!sb.isEmpty()) {
            sb.append('\n');
        }
        sb.append("機械名")
                .append('\t')
                .append("工程名")
                .append('\t')
                .append("日付")
                .append('\t')
                .append("計画m")
                .append('\t')
                .append("日報完了");
        for (PlanRow row : rows) {
            sb.append('\n');
            appendTsvCell(sb, row.getMachineName());
            sb.append('\t');
            appendTsvCell(sb, row.getProcessName());
            sb.append('\t');
            appendTsvCell(sb, row.getDateYmd());
            sb.append('\t');
            appendTsvCell(sb, row.getPlanMeters());
            sb.append('\t');
            appendTsvCell(sb, row.getDailyReportCompletion());
        }
        return sb.toString();
    }

    private static void appendPlanCopyHeaderTsv(StringBuilder sb, PlanCopyHeader header) {
        if (header == null) {
            return;
        }
        appendHeaderTsvLineRequired(sb, "依頼No", header.iraiNo());
        appendHeaderTsvLineRequired(sb, "契約NO", header.contractNo());
        appendHeaderTsvLineRequired(sb, "原反投入日", header.rawInputDate());
    }

    private static void appendHeaderTsvLineRequired(StringBuilder sb, String label, String value) {
        if (!sb.isEmpty()) {
            sb.append('\n');
        }
        appendTsvCell(sb, label);
        sb.append('\t');
        appendTsvCell(sb, value != null ? value : "");
    }

    private static void appendPlanCopyHeaderHtml(StringBuilder sb, PlanCopyHeader header) {
        if (header == null) {
            return;
        }
        List<String[]> lines = headerLines(header);
        if (lines.isEmpty()) {
            return;
        }
        sb.append(
                "<table border=\"0\" cellspacing=\"0\" cellpadding=\"2\""
                        + " style=\"border-collapse:collapse;font-family:'Meiryo UI',sans-serif;font-size:11pt;margin-bottom:8px;\">");
        for (String[] line : lines) {
            sb.append("<tr><td style=\"padding:2px 12px 2px 0;font-weight:bold;white-space:nowrap;\">")
                    .append(ClipboardTableSupport.escapeHtml(line[0]))
                    .append("</td><td style=\"padding:2px 0;\">")
                    .append(ClipboardTableSupport.escapeHtml(line[1]))
                    .append("</td></tr>");
        }
        sb.append("</table>");
    }

    private static List<String[]> headerLines(PlanCopyHeader header) {
        List<String[]> lines = new ArrayList<>();
        lines.add(new String[] {"依頼No", header.iraiNo() != null ? header.iraiNo() : ""});
        lines.add(new String[] {"契約NO", header.contractNo() != null ? header.contractNo() : ""});
        lines.add(
                new String[] {
                    "原反投入日", header.rawInputDate() != null ? header.rawInputDate() : ""
                });
        return lines;
    }

    private static void appendTsvCell(StringBuilder sb, String value) {
        appendDelimitedCell(sb, value, '\t');
    }

    static void appendCsvCell(StringBuilder sb, String value) {
        appendDelimitedCell(sb, value, ',');
    }

    private static void appendDelimitedCell(StringBuilder sb, String value, char delimiter) {
        String text = value != null ? value : "";
        if (text.indexOf(delimiter) >= 0
                || text.indexOf('"') >= 0
                || text.indexOf('\t') >= 0
                || text.indexOf('\n') >= 0
                || text.indexOf('\r') >= 0) {
            sb.append('"').append(text.replace("\"", "\"\"")).append('"');
        } else {
            sb.append(text);
        }
    }

    private void applyFilter() {
        String q =
                filterField.getText() != null
                        ? filterField.getText().strip().toLowerCase(Locale.ROOT)
                        : "";
        boolean hideNoOriginal = hideNoOriginalCheck != null && hideNoOriginalCheck.isSelected();
        boolean hideOldJuchuInput =
                hideOldJuchuInputCheck != null && hideOldJuchuInputCheck.isSelected();
        int juchuInputHideDays = hideOldJuchuInput ? resolveJuchuInputHideDays() : 0;
        boolean showNoAladdinOnly =
                showNoAladdinCheck != null && showNoAladdinCheck.isSelected();
        boolean showAdjustDeliveryFromTodayOnly =
                showAdjustDeliveryFromTodayCheck != null
                        && showAdjustDeliveryFromTodayCheck.isSelected();
        boolean showDailyReportIncompleteOnly =
                showDailyReportIncompleteOnlyCheck == null
                        || showDailyReportIncompleteOnlyCheck.isSelected();

        filteredRows.setPredicate(
                row -> {
                    PipelineStatusRow src = row.source();
                    if (src == null) {
                        return false;
                    }
                    if (hideNoOriginal && !src.originalPresent()) {
                        return false;
                    }
                    if (hideOldJuchuInput
                            && RequestFormPipelineStatusService.shouldHideByJuchuInputDate(
                                    src.juchuInputDate(), juchuInputHideDays)) {
                        return false;
                    }
                    if (showNoAladdinOnly && aladdinJsonAvailable && src.aladdinPresent()) {
                        return false;
                    }
                    if (showAdjustDeliveryFromTodayOnly
                            && RequestFormPipelineStatusService
                                    .shouldHideByAdjustDeliveryBeforeToday(
                                            src.juchuAdjustDeliveryDate())) {
                        return false;
                    }
                    if (showDailyReportIncompleteOnly
                            && "完了".equals(nullToEmpty(row.getDailyReportOrderStatus()))) {
                        return false;
                    }
                    if (q.isEmpty()) {
                        return true;
                    }
                    String irai =
                            row.getIraiNo() != null ? row.getIraiNo().toLowerCase(Locale.ROOT) : "";
                    String user =
                            src.user() != null ? src.user().toLowerCase(Locale.ROOT) : "";
                    return irai.contains(q) || user.contains(q);
                });
        updateStatusLabel();
    }

    private int resolveJuchuInputHideDays() {
        if (juchuInputHideDaysSpinner == null || juchuInputHideDaysSpinner.getValue() == null) {
            return RequestFormPipelineStatusService.DEFAULT_JUCHU_INPUT_DATE_HIDE_DAYS;
        }
        return Math.max(0, juchuInputHideDaysSpinner.getValue());
    }

    private int countHiddenByJuchuInputDate() {
        if (hideOldJuchuInputCheck == null || !hideOldJuchuInputCheck.isSelected()) {
            return 0;
        }
        int days = resolveJuchuInputHideDays();
        int count = 0;
        for (MainRow row : allRows) {
            PipelineStatusRow src = row.source();
            if (src != null
                    && RequestFormPipelineStatusService.shouldHideByJuchuInputDate(
                            src.juchuInputDate(), days)) {
                count++;
            }
        }
        return count;
    }

    private int countHiddenByNoAladdinOnly() {
        if (!aladdinJsonAvailable
                || showNoAladdinCheck == null
                || !showNoAladdinCheck.isSelected()) {
            return 0;
        }
        int count = 0;
        for (MainRow row : allRows) {
            PipelineStatusRow src = row.source();
            if (src != null && src.aladdinPresent()) {
                count++;
            }
        }
        return count;
    }

    private int countHiddenByAdjustDeliveryBeforeToday() {
        if (showAdjustDeliveryFromTodayCheck == null
                || !showAdjustDeliveryFromTodayCheck.isSelected()) {
            return 0;
        }
        int count = 0;
        for (MainRow row : allRows) {
            PipelineStatusRow src = row.source();
            if (src != null
                    && RequestFormPipelineStatusService.shouldHideByAdjustDeliveryBeforeToday(
                            src.juchuAdjustDeliveryDate())) {
                count++;
            }
        }
        return count;
    }

    private int countHiddenByDailyReportComplete() {
        if (showDailyReportIncompleteOnlyCheck != null
                && !showDailyReportIncompleteOnlyCheck.isSelected()) {
            return 0;
        }
        int count = 0;
        for (MainRow row : allRows) {
            if ("完了".equals(nullToEmpty(row.getDailyReportOrderStatus()))) {
                count++;
            }
        }
        return count;
    }

    private void updateStatusLabel() {
        int visible = filteredRows.size();
        int total = allRows.size();
        StringBuilder status = new StringBuilder();
        status.append("表示: ").append(visible).append(" / 全 ").append(total).append(" 件");
        int hiddenByInputDate = countHiddenByJuchuInputDate();
        if (hiddenByInputDate > 0) {
            status.append(" | 非表示: ")
                    .append(hiddenByInputDate)
                    .append("件（受注入力日")
                    .append(resolveJuchuInputHideDays())
                    .append("日以上前）");
        }
        int hiddenByNoAladdin = countHiddenByNoAladdinOnly();
        if (hiddenByNoAladdin > 0) {
            status.append(" | 非表示: ")
                    .append(hiddenByNoAladdin)
                    .append("件（アラジン計画あり）");
        }
        int hiddenByAdjustDelivery = countHiddenByAdjustDeliveryBeforeToday();
        if (hiddenByAdjustDelivery > 0) {
            status.append(" | 非表示: ")
                    .append(hiddenByAdjustDelivery)
                    .append("件（調整納期が当日より前または未設定）");
        }
        int hiddenByDailyReportComplete = countHiddenByDailyReportComplete();
        if (hiddenByDailyReportComplete > 0) {
            status.append(" | 非表示: ")
                    .append(hiddenByDailyReportComplete)
                    .append("件（日報完了）");
        }
        if (aladdinJsonAvailable) {
            status.append(" | ①〜⑦=依頼ごとの計画日（昇順・最大7日）");
        }
        if (!lastScanWarnings.isEmpty()) {
            status.append(" | ").append(lastScanWarnings);
        }
        statusLabel.setText(status.toString());
    }

    private void updateDailyReportSourceLabel() {
        if (dailyReportSourceLabel == null) {
            return;
        }
        String path = dailyReportLookup != null ? dailyReportLookup.sourcePath() : "";
        if (path.isBlank()) {
            dailyReportSourceLabel.setText("加工日報: （未読込）");
            dailyReportSourceLabel.setManaged(true);
            dailyReportSourceLabel.setVisible(true);
            return;
        }
        dailyReportSourceLabel.setText("加工日報: " + path);
        dailyReportSourceLabel.setManaged(true);
        dailyReportSourceLabel.setVisible(true);
    }

    private List<String> mainTableVisibleColumnTitles() {
        if (mainTable == null) {
            return mainTableColumnTitles();
        }
        List<String> titles = new ArrayList<>();
        for (TableColumn<MainRow, ?> column : mainTable.getColumns()) {
            if (column.isVisible()) {
                titles.add(column.getText() != null ? column.getText() : "");
            }
        }
        return titles.isEmpty() ? mainTableColumnTitles() : List.copyOf(titles);
    }

    private void setupMainColumnsOnce() {
        if (mainColumnsInitialized || mainTable == null) {
            return;
        }
        mainColumnsInitialized = true;
        mainColumnTitleToProperty.clear();
        for (MainColDef def : mainColumnDefs) {
            mainColumnTitleToProperty.put(def.title(), def.property());
        }

        List<TableColumn<MainRow, String>> columns = new ArrayList<>();
        for (MainColDef def : mainColumnDefs) {
            TableColumn<MainRow, String> column = new TableColumn<>(def.title());
            column.setCellValueFactory(new PropertyValueFactory<>(def.property()));
            column.setMinWidth(def.defaultWidth() * 0.6);
            column.setPrefWidth(def.defaultWidth());
            column.setReorderable(true);
            column.setUserData(def.property());
            columns.add(column);
        }
        suppressColumnPersistence.set(true);
        try {
            mainTable.getColumns().setAll(columns);
            mainTable.setColumnResizePolicy(TableView.UNCONSTRAINED_RESIZE_POLICY);
            List<TableColumnOrderPersistence.ColumnSpec> layout =
                    TableColumnOrderPersistence.loadLayout(
                            TableColumnOrderPersistence.TableId.REQUEST_FORM_PIPELINE_CHECK);
            if (!layout.isEmpty()) {
                TableColumnOrderPersistence.applyOrderToTableColumns(
                        mainTable,
                        layout.stream()
                                .map(TableColumnOrderPersistence.ColumnSpec::title)
                                .toList());
                TableColumnOrderPersistence.applyWidthsToTableColumns(mainTable, layout, 112);
            }
        } finally {
            suppressColumnPersistence.set(false);
        }

        TableColumnOrderPersistence.installColumnLayoutWatcher(
                mainTable,
                TableColumnOrderPersistence.TableId.REQUEST_FORM_PIPELINE_CHECK,
                suppressColumnPersistence::get);

        Runnable resetMainColumns =
                () -> {
                    suppressColumnPersistence.set(true);
                    try {
                        for (int i = 0; i < columns.size() && i < mainColumnDefs.size(); i++) {
                            double w = mainColumnDefs.get(i).defaultWidth();
                            columns.get(i).setMinWidth(w * 0.6);
                            columns.get(i).setPrefWidth(w);
                        }
                    } finally {
                        suppressColumnPersistence.set(false);
                    }
                };

        Window owner = shell != null ? shell.getPrimaryStage() : null;
        if (mainColumnStripHost != null) {
            mainColumnStripHost
                    .getChildren()
                    .setAll(
                            TableViewColumnSettingsStrip.create(
                                    mainTable,
                                    resetMainColumns,
                                    false,
                                    TableColumnOrderPersistence.TableId.REQUEST_FORM_PIPELINE_CHECK,
                                    null,
                                    () ->
                                            ColumnVisibilitySupport
                                                    .openTableViewColumnVisibilityDialog(
                                                            owner,
                                                            TableColumnOrderPersistence.TableId
                                                                    .REQUEST_FORM_PIPELINE_CHECK,
                                                            mainTable)));
        }

        Platform.runLater(
                () ->
                        ColumnVisibilitySupport.applyColumnVisibilityToTableView(
                                mainTable,
                                TableColumnOrderPersistence.loadColumnVisibility(
                                        TableColumnOrderPersistence.TableId
                                                .REQUEST_FORM_PIPELINE_CHECK,
                                        mainTable.getColumns().size())));
    }

    private void setupMismatchColumns() {
        mismatchTable.getColumns().setAll(
                colMismatch("フォーム項目", "formLabel", 220),
                colMismatch("原本値", "originalValue", 160),
                colMismatch("受注ファイル値", "juchuValue", 160));
    }

    private void setupPlanColumns() {
        planTable.getColumns().setAll(
                colPlan("機械名", "machineName", 140),
                colPlan("工程名", "processName", 100),
                colPlan("日付", "dateYmd", 100),
                colPlan("計画m", "planMeters", 80),
                colPlan("日報完了", "dailyReportCompletion", 72));
    }

    private static TableColumn<MismatchRow, String> colMismatch(
            String title, String prop, double width) {
        TableColumn<MismatchRow, String> c = new TableColumn<>(title);
        c.setCellValueFactory(new PropertyValueFactory<>(prop));
        c.setPrefWidth(width);
        return c;
    }

    private static TableColumn<PlanRow, String> colPlan(String title, String prop, double width) {
        TableColumn<PlanRow, String> c = new TableColumn<>(title);
        c.setCellValueFactory(new PropertyValueFactory<>(prop));
        c.setPrefWidth(width);
        return c;
    }
}
