package jp.co.pm.ai.desktop;

import java.awt.Desktop;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.function.Consumer;

import javafx.animation.KeyFrame;
import javafx.animation.Timeline;
import javafx.application.Platform;
import javafx.beans.binding.Bindings;
import javafx.beans.property.BooleanProperty;
import javafx.beans.property.SimpleBooleanProperty;
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
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.Window;
import javafx.util.Duration;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.NetworkSourceDirResolver;
import jp.co.pm.ai.desktop.config.SourceFileExtensionPolicy;

import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup;
import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup.PlanEntry;
import jp.co.pm.ai.desktop.ui.ButtonAttentionGlow;
import jp.co.pm.ai.desktop.ui.ClipboardTableSupport;
import jp.co.pm.ai.desktop.ui.ColumnVisibilitySupport;
import jp.co.pm.ai.desktop.ui.SourceExtensionErrorOverlay;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;
import jp.co.pm.ai.desktop.ui.TableViewColumnSettingsStrip;
import jp.co.pm.ai.desktop.reconciliation.JuchuTransferCoverageCheck.ColumnCheck;
import jp.co.pm.ai.desktop.reconciliation.JuchuHeaderAliasRegistry;
import jp.co.pm.ai.desktop.reconciliation.JuchuTransferValueNormalizer;
import jp.co.pm.ai.desktop.reconciliation.KonanDailyReportLookup;
import jp.co.pm.ai.desktop.reconciliation.RawInputDateCrossSourceCheck;
import jp.co.pm.ai.desktop.reconciliation.RemoteDesktopLatestSourceFiles;
import jp.co.pm.ai.desktop.io.AladdinProcessingPlanSourceReloader;
import jp.co.pm.ai.desktop.reconciliation.RequestFormPipelineIssueCheck;
import jp.co.pm.ai.desktop.reconciliation.RequestFormPipelineIssueCheck.IssueKind;
import jp.co.pm.ai.desktop.reconciliation.RequestFormPipelineStatusService;
import jp.co.pm.ai.desktop.reconciliation.RequestFormPipelineStatusService.PipelineStatusRow;
import jp.co.pm.ai.desktop.reconciliation.RequestFormPipelineStatusService.ScanResult;

/** 依頼書原本の受注転記率・アラジン加工計画の有無を一覧表示する。 */
public final class RequestFormPipelineCheckTabController {

    private static final int PLAN_DAY_COLUMNS =
            AladdinShapedPlanQtyLookup.PIPELINE_CHECK_PLAN_DAY_COLUMNS;

    private static final String DAILY_REPORT_ORDER_STATUS_ABSENT = "―";
    private static final String DAILY_REPORT_ORDER_STATUS_COMPLETE = "完了";

    private static final Duration ALADDIN_PLAN_WATCH_INTERVAL = Duration.minutes(1);
    private static final int APPLY_SCAN_RESULT_BATCH_SIZE = 50;
    private static final long PIPELINE_PROGRESS_UI_INTERVAL_MS = 300L;

    private static final String ORIGINAL_DIR_UNCONFIGURED_STATUS =
            "依頼書原本フォルダ未設定: 環境変数タブで "
                    + AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR
                    + " を設定してください（走査は実行しません）";

    private static final String HINT_TEXT =
            "「更新」…加工計画の再読込・原本走査・転記率と shaped_aladdin_plan.json を照合。"
                    + " 使用中の加工計画・加工日報がリモートデスクトップの取得データ最新と異なると更新ボタンが点灯。\n"
                    + "・依頼書原本フォルダ（環境変数 "
                    + AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR
                    + "）を設定するまで走査しません\n"
                    + "・①〜⑦ … 依頼ごとの計画日（昇順・最大7日。例: 7/3 100m）\n"
                    + "・投入日一致 … 原反投入日を4ソース照合（全一致のみ「一致」）\n"
                    + "・確認 … 要チェック行はフィルタ無視で表示（日報「完了」は不要）。一括チェック可\n"
                    + "・段階1 … 未確認がある間は実行不可。計画更新で確認リセット（1分監視）。"
                    + " 加工計画最新が.xlsx以外／加工日報最新が.csv以外のときも実行不可\n"
                    + "・依頼NO先頭「2」… 自社加工品（配台対象外・段階1要確認対象外）";

    private record MainColDef(String title, String property, double defaultWidth) {}

    private static List<MainColDef> defaultMainColumnDefs() {
        List<MainColDef> defs = new ArrayList<>();
        defs.add(new MainColDef("依頼No", "iraiNo", 90));
        defs.add(new MainColDef("問題", "issueSummary", 168));
        defs.add(new MainColDef("確認", "issueConfirmedDisplay", 52));
        defs.add(new MainColDef("原本", "originalFile", 140));
        defs.add(new MainColDef("回答日", "indexResponseDate", 88));
        defs.add(new MainColDef("原反投入日", "indexInputDate", 88));
        defs.add(new MainColDef("投入日一致", "rawInputDateMatchStatus", 72));
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
        /** 原反投入日4ソース照合ステータス: 一致 / 不一致 / ― */
        private String rawInputDateMatchStatus;
        /** 加工日報の依頼単位ステータス: 完了 / 未了 / ― */
        private String dailyReportOrderStatus;
        private String issueSummary = "";
        private boolean hasIssues;
        private final BooleanProperty issueConfirmed = new SimpleBooleanProperty(false);
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

        public String getRawInputDateMatchStatus() {
            return rawInputDateMatchStatus;
        }

        public void setRawInputDateMatchStatus(String rawInputDateMatchStatus) {
            this.rawInputDateMatchStatus = rawInputDateMatchStatus;
        }

        public String getDailyReportOrderStatus() {
            return dailyReportOrderStatus;
        }

        public void setDailyReportOrderStatus(String dailyReportOrderStatus) {
            this.dailyReportOrderStatus = dailyReportOrderStatus;
        }

        public String getIssueSummary() {
            return issueSummary;
        }

        public void setIssueSummary(String issueSummary) {
            this.issueSummary = issueSummary != null ? issueSummary : "";
        }

        public boolean hasIssues() {
            return hasIssues;
        }

        public void setHasIssues(boolean hasIssues) {
            this.hasIssues = hasIssues;
        }

        public BooleanProperty issueConfirmedProperty() {
            return issueConfirmed;
        }

        public boolean isIssueConfirmed() {
            return issueConfirmed.get();
        }

        public String getIssueConfirmedDisplay() {
            if (!RequestFormPipelineCheckTabController.requiresStage1Confirmation(this)) {
                return "―";
            }
            return RequestFormPipelineIssueCheck.formatConfirmedDisplay(true, issueConfirmed.get());
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

    /** 下段「原反投入日 4ソース照合」表の1行（ソース・値・状態）。 */
    public static final class CrossSourceRow {
        private final SimpleStringProperty source = new SimpleStringProperty();
        private final SimpleStringProperty value = new SimpleStringProperty();
        private final SimpleStringProperty status = new SimpleStringProperty();

        CrossSourceRow(String source, String value, String status) {
            this.source.set(source != null ? source : "");
            this.value.set(value != null ? value : "");
            this.status.set(status != null ? status : "");
        }

        public String getSource() {
            return source.get();
        }

        public SimpleStringProperty sourceProperty() {
            return source;
        }

        public String getValue() {
            return value.get();
        }

        public SimpleStringProperty valueProperty() {
            return value;
        }

        public String getStatus() {
            return status.get();
        }

        public SimpleStringProperty statusProperty() {
            return status;
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
    private CheckBox showRawInputMismatchOnlyCheck;

    @FXML
    private Label statusLabel;

    @FXML
    private Label hintLabel;

    @FXML
    private Label stage1GateLabel;

    @FXML
    private Label aladdinPlanSourceLabel;

    @FXML
    private Button openAladdinPlanSourceFolderButton;

    @FXML
    private Label dailyReportSourceLabel;

    @FXML
    private Button openDailyReportSourceFolderButton;

    @FXML
    private TableView<MainRow> mainTable;

    @FXML
    private StackPane mainTableHost;

    @FXML
    private Button copyMainTableButton;

    @FXML
    private Button copyMainTableForEmailButton;

    @FXML
    private Button copyMainTableHtmlButton;

    @FXML
    private Button copyMainTableCsvButton;

    @FXML
    private Button confirmAllVisibleIssuesButton;

    @FXML
    private HBox mainColumnStripHost;

    @FXML
    private TableView<CrossSourceRow> crossSourceTable;

    @FXML
    private Button copyCrossSourceTableButton;

    @FXML
    private Button copyCrossSourceTableForEmailButton;

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
    private final ObservableList<CrossSourceRow> crossSourceRows = FXCollections.observableArrayList();
    private final ObservableList<PlanRow> planRows = FXCollections.observableArrayList();

    /** 下段アラジン計画表の表示元（コピー時の依頼ヘッダ）。 */
    private MainRow planContextRow;

    private KonanDailyReportLookup dailyReportLookup = KonanDailyReportLookup.empty();

    /** 直近の「更新」で読み込んだアラジン加工計画ソースの絶対パス。 */
    private String aladdinPlanSourcePath = "";

    /**
     * 直近の走査完了時点のアラジン加工計画ソース版。
     * ディスク上の最新版と異なれば再走査が必要。
     */
    private String lastScannedAladdinPlanSourceRevisionKey = "";

    /**
     * 直近に取り込んだアラジン加工計画ソースの版（絶対パス + 最終更新時刻）。
     * 変更検知で確認チェックをリセットする。
     */
    private String lastAladdinPlanSourceRevisionKey = "";

    private ButtonAttentionGlow refreshButtonGlow;
    private Timeline aladdinPlanWatchTimeline;

    private boolean aladdinJsonAvailable = true;
    private String lastScanWarnings = "";
    /** 起動後に一度でも走査結果を反映したら true（手動「更新」は常に可）。 */
    private boolean scanApplied;
    private boolean refreshInProgress;
    private volatile Consumer<Boolean> refreshCompleteCallback;
    private volatile long lastPipelineProgressUiAt;
    private volatile String latestPipelineProgressDetail = "";
    private volatile double latestPipelineProgressFraction = Double.NaN;

    private final List<MainColDef> mainColumnDefs = new ArrayList<>(defaultMainColumnDefs());

    private final Map<String, String> mainColumnTitleToProperty = new LinkedHashMap<>();

    private final AtomicBoolean suppressColumnPersistence = new AtomicBoolean(false);

    private boolean mainColumnsInitialized;

    @FXML
    private void initialize() {
        hintLabel.setText(HINT_TEXT);
        updateAladdinPlanSourceLabel();
        updateDailyReportSourceLabel();
        filteredRows = new FilteredList<>(allRows, row -> true);
        mainTable.setItems(filteredRows);
        mismatchTable.setItems(mismatchRows);
        mismatchTable.setPlaceholder(new Label("相違なし"));
        crossSourceTable.setItems(crossSourceRows);
        crossSourceTable.setPlaceholder(new Label("行を選択すると原反投入日の照合結果を表示"));
        planTable.setItems(planRows);
        installMainTableScrollLayout();

        setupCrossSourceColumns();
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
        if (confirmAllVisibleIssuesButton != null) {
            confirmAllVisibleIssuesButton.setDisable(true);
        }
        if (copyPlanTableButton != null) {
            copyPlanTableButton.disableProperty().bind(Bindings.isEmpty(planRows));
        }
        if (copyPlanTableForEmailButton != null) {
            copyPlanTableForEmailButton.disableProperty().bind(Bindings.isEmpty(planRows));
        }
        if (copyCrossSourceTableButton != null) {
            copyCrossSourceTableButton.disableProperty().bind(Bindings.isEmpty(crossSourceRows));
        }
        if (copyCrossSourceTableForEmailButton != null) {
            copyCrossSourceTableForEmailButton.disableProperty()
                    .bind(Bindings.isEmpty(crossSourceRows));
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
        if (showRawInputMismatchOnlyCheck != null) {
            showRawInputMismatchOnlyCheck
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
                    .addListener((obs, o, n) -> onJuchuInputHideScanSettingChanged());
        }
        if (hideOldJuchuInputCheck != null) {
            hideOldJuchuInputCheck
                    .selectedProperty()
                    .addListener(
                            (obs, o, n) -> {
                                if (juchuInputHideDaysSpinner != null) {
                                    juchuInputHideDaysSpinner.setDisable(!n);
                                }
                                onJuchuInputHideScanSettingChanged();
                            });
            if (juchuInputHideDaysSpinner != null) {
                juchuInputHideDaysSpinner.setDisable(!hideOldJuchuInputCheck.isSelected());
            }
        }

        mainTable.getSelectionModel()
                .selectedItemProperty()
                .addListener((obs, oldRow, newRow) -> showDetail(newRow));
    }

    /** 親幅に収め、列合計が広いときは TableView 内の横スクロールを有効にする。 */
    private void installMainTableScrollLayout() {
        if (mainTable == null || mainTableHost == null) {
            return;
        }
        VBox.setVgrow(mainTableHost, Priority.ALWAYS);
        mainTable.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
        mainTable.prefWidthProperty().bind(mainTableHost.widthProperty());
        mainTable.prefHeightProperty().bind(mainTableHost.heightProperty());
        mainTable.setColumnResizePolicy(TableView.UNCONSTRAINED_RESIZE_POLICY);
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        setupMainColumnsOnce();
        if (refreshButton != null && refreshButtonGlow == null) {
            refreshButtonGlow = new ButtonAttentionGlow(refreshButton);
        }
        startAladdinPlanWatch();
        updateStage1GateLabel();
        notifyStage1GateChanged();
        Platform.runLater(this::tickAladdinPlanWatch);
    }

    /** 工場切替後: 走査結果を破棄し、タブ選択時まで再走査を遅延する。 */
    void onFactorySiteChanged(boolean lightweight) {
        scanApplied = false;
        lastScanWarnings = "";
        aladdinPlanSourcePath = "";
        lastScannedAladdinPlanSourceRevisionKey = "";
        lastAladdinPlanSourceRevisionKey = "";
        dailyReportLookup = KonanDailyReportLookup.empty();
        stopRefreshButtonGlow();
        allRows.clear();
        mismatchRows.clear();
        crossSourceRows.clear();
        planRows.clear();
        planContextRow = null;
        if (statusLabel != null) {
            statusLabel.setText("工場切替: タブを開くと再走査します");
        }
        updateAladdinPlanSourceLabel();
        updateDailyReportSourceLabel();
        updateStage1GateLabel();
        notifyStage1GateChanged();
    }

    /** メインシェルで当該タブが選択されたとき。未走査または加工計画更新後は自動走査する。 */
    void onMainShellTabSelected() {
        scheduleRefreshIfNeededOnTabSelected();
    }

    /** 依頼書原本フォルダが環境変数タブで設定されたあと、必要なら走査を予約する。 */
    void onRequestFormOriginalDirEnvConfigured() {
        scanApplied = false;
        scheduleRefreshIfNeededOnTabSelected();
    }

    /** 受注ファイルパスが再設定されたあと、転記率照合などを再スキャンする。 */
    void onRequestFormJuchuFileChanged() {
        scanApplied = false;
        scheduleRefreshIfNeededOnTabSelected();
    }

    private boolean isRequestFormOriginalDirEnvConfigured() {
        return shell != null
                && AppPaths.isRequestFormOriginalDirEnvConfigured(shell.snapshotUiEnv());
    }

    private void applyPendingOriginalDirSetupState() {
        scanApplied = false;
        allRows.clear();
        mismatchRows.clear();
        crossSourceRows.clear();
        planRows.clear();
        planContextRow = null;
        lastScanWarnings = "";
        if (statusLabel != null) {
            statusLabel.setText(ORIGINAL_DIR_UNCONFIGURED_STATUS);
        }
        updateStage1GateLabel();
        notifyStage1GateChanged();
    }

    private boolean ensureRequestFormOriginalDirConfiguredForScan() {
        if (isRequestFormOriginalDirEnvConfigured()) {
            return true;
        }
        applyPendingOriginalDirSetupState();
        return false;
    }

    private void scheduleRefreshIfNeededOnTabSelected() {
        if (refreshInProgress || shell == null || shell.isFactorySiteSwitchInProgress()) {
            return;
        }
        if (!isRequestFormOriginalDirEnvConfigured()) {
            Platform.runLater(this::applyPendingOriginalDirSetupState);
            return;
        }
        if (!scanApplied) {
            Platform.runLater(() -> startRefresh(true));
            return;
        }
        if (isAladdinPlanSourceNewerThanLastScan(shell.snapshotUiEnv())) {
            Platform.runLater(() -> startRefresh(true));
        } else {
            Platform.runLater(this::refreshAladdinPlanWatchState);
        }
    }

    /** 起動後バックグラウンド読込（MainShell コーディネータから呼ぶ）。 */
    void preloadInBackground(Consumer<Boolean> onComplete) {
        if (shell == null) {
            completeRefreshPreload(false, onComplete);
            return;
        }
        if (shell.isFactorySiteSwitchInProgress()) {
            completeRefreshPreload(false, onComplete);
            return;
        }
        if (!isRequestFormOriginalDirEnvConfigured()) {
            Platform.runLater(this::applyPendingOriginalDirSetupState);
            completeRefreshPreload(false, onComplete);
            return;
        }
        if (refreshInProgress) {
            completeRefreshPreload(false, onComplete);
            return;
        }
        if (scanApplied && !isAladdinPlanSourceNewerThanLastScan(shell.snapshotUiEnv())) {
            completeRefreshPreload(true, onComplete);
            return;
        }
        refreshCompleteCallback = onComplete;
        startRefresh(true);
    }

    private void completeRefreshPreload(boolean ok, Consumer<Boolean> onComplete) {
        clearPipelineCheckProgress();
        Consumer<Boolean> pending = onComplete != null ? onComplete : refreshCompleteCallback;
        refreshCompleteCallback = null;
        if (pending == null) {
            return;
        }
        Platform.runLater(() -> pending.accept(ok));
    }

    private void scheduleInitialRefreshIfNeeded() {
        scheduleRefreshIfNeededOnTabSelected();
    }

    @FXML
    private void onRefreshButtonAction() {
        startRefresh(false);
    }

    private void startRefresh() {
        startRefresh(false);
    }

    private void startRefresh(boolean skipStaleSourceConfirm) {
        if (!skipStaleSourceConfirm && !confirmRefreshDespiteStaleSources()) {
            return;
        }
        stopRefreshButtonGlow();
        if (shell == null) {
            statusLabel.setText("シェル未接続");
            return;
        }
        if (refreshInProgress) {
            return;
        }
        if (!ensureRequestFormOriginalDirConfiguredForScan()) {
            completeRefreshPreload(false, null);
            return;
        }
        refreshInProgress = true;
        setRefreshing(true);
        statusLabel.setText("アラジン加工計画読込中…");
        reportPipelineCheckProgress(Double.NaN, "アラジン加工計画読込中…");
        JuchuHeaderAliasRegistry registry = shell.snapshotJuchuHeaderAliasRegistryForExport();
        RequestFormPipelineStatusService.ScanProgressListener scanProgress =
                (fraction, detail) -> reportPipelineCheckProgress(fraction, detail);
        Thread worker =
                new Thread(
                        () -> {
                            Map<String, String> ui = shell.snapshotUiEnv();
                            List<String> reloadWarnings = new ArrayList<>();
                            String loadedPlanSourcePath = aladdinPlanSourcePath;
                            boolean aladdinPlanSourceUpdated = false;
                            try {
                                try {
                                    AladdinProcessingPlanSourceReloader.ReloadResult reload =
                                            AladdinProcessingPlanSourceReloader
                                                    .reloadNewestFromDiskAndSaveShapedJson(ui);
                                    loadedPlanSourcePath =
                                            reload.sourceFile().toAbsolutePath().normalize().toString();
                                    aladdinPlanSourceUpdated =
                                            registerAladdinPlanSourceReload(reload.sourceFile());
                                    shell.appendLog(
                                            "[pipeline-check] アラジン加工計画再読込: "
                                                    + reload.sourceFile()
                                                    + " ("
                                                    + reload.rowCount()
                                                    + " 行 × "
                                                    + reload.columnCount()
                                                    + " 列)");
                                    Platform.runLater(shell::refreshAladdinProcessingPlanTabFromDisk);
                                } catch (Exception reloadEx) {
                                    String msg =
                                            "アラジン加工計画の再読込に失敗: "
                                                    + (reloadEx.getMessage() != null
                                                            ? reloadEx.getMessage()
                                                            : reloadEx.toString());
                                    reloadWarnings.add(msg);
                                    shell.appendLog("[pipeline-check] " + msg);
                                }
                                Platform.runLater(() -> statusLabel.setText("走査中…"));
                                int juchuInputHideDays = resolveJuchuInputHideDaysForScan();
                                ScanResult result =
                                        RequestFormPipelineStatusService.scan(
                                                ui, registry, scanProgress, juchuInputHideDays);
                                if (!reloadWarnings.isEmpty()) {
                                    List<String> mergedWarnings = new ArrayList<>(reloadWarnings);
                                    mergedWarnings.addAll(result.warnings());
                                    result =
                                            new ScanResult(
                                                    result.rows(),
                                                    mergedWarnings,
                                                    result.aladdinJsonAvailable(),
                                                    result.planDateHeaders(),
                                                    result.dailyReportLookup());
                                }
                                ScanResult scanResult = result;
                                String planSourcePath = loadedPlanSourcePath;
                                boolean resetConfirmations = aladdinPlanSourceUpdated;
                                Platform.runLater(
                                        () -> {
                                            setRefreshing(false);
                                            aladdinPlanSourcePath = planSourcePath;
                                            applyScanResult(scanResult, resetConfirmations);
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
                                            completeRefreshPreload(false, null);
                                        });
                            }
                        },
                        "request-form-pipeline-check");
        worker.setDaemon(true);
        worker.setPriority(Thread.MIN_PRIORITY);
        worker.start();
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

    private void reportPipelineCheckProgress(double fraction, String detail) {
        if (shell == null) {
            return;
        }
        latestPipelineProgressFraction = fraction;
        latestPipelineProgressDetail = detail != null ? detail : "";
        long now = System.currentTimeMillis();
        boolean force = !Double.isNaN(fraction) && (fraction <= 0.02 || fraction >= 0.99);
        if (!force && now - lastPipelineProgressUiAt < PIPELINE_PROGRESS_UI_INTERVAL_MS) {
            return;
        }
        lastPipelineProgressUiAt = now;
        final double fractionToShow = latestPipelineProgressFraction;
        final String detailToShow = latestPipelineProgressDetail;
        Platform.runLater(() -> shell.setGlobalLongTaskProgress(fractionToShow, detailToShow));
    }

    private void flushPipelineCheckProgressUi() {
        if (shell == null) {
            return;
        }
        final double fractionToShow = latestPipelineProgressFraction;
        final String detailToShow = latestPipelineProgressDetail;
        if (Platform.isFxApplicationThread()) {
            shell.setGlobalLongTaskProgress(fractionToShow, detailToShow);
        } else {
            Platform.runLater(() -> shell.setGlobalLongTaskProgress(fractionToShow, detailToShow));
        }
    }

    private void clearPipelineCheckProgress() {
        if (shell == null) {
            return;
        }
        Platform.runLater(shell::clearGlobalLongTaskProgress);
    }

    private void applyScanResult(ScanResult result) {
        applyScanResult(result, true);
    }

    private void applyScanResult(ScanResult result, boolean resetIssueConfirmations) {
        Map<String, Boolean> previousConfirmations = new HashMap<>();
        if (!resetIssueConfirmations) {
            for (MainRow row : allRows) {
                if (row.isIssueConfirmed() && row.getIraiNo() != null && !row.getIraiNo().isBlank()) {
                    previousConfirmations.put(row.getIraiNo(), true);
                }
            }
        }
        scanApplied = true;
        allRows.clear();
        aladdinJsonAvailable = result.aladdinJsonAvailable();
        dailyReportLookup =
                result.dailyReportLookup() != null
                        ? result.dailyReportLookup()
                        : KonanDailyReportLookup.empty();
        if (statusLabel != null) {
            statusLabel.setText("一覧を反映中…");
        }
        applyScanResultBatch(result, resetIssueConfirmations, previousConfirmations, 0);
    }

    private void applyScanResultBatch(
            ScanResult result,
            boolean resetIssueConfirmations,
            Map<String, Boolean> previousConfirmations,
            int startIndex) {
        List<PipelineStatusRow> sourceRows = result.rows();
        int endIndex =
                Math.min(startIndex + APPLY_SCAN_RESULT_BATCH_SIZE, sourceRows.size());
        for (int i = startIndex; i < endIndex; i++) {
            PipelineStatusRow row = sourceRows.get(i);
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
            ui.setRawInputDateMatchStatus(
                    row.rawInputDateCrossCheck() != null
                            ? row.rawInputDateCrossCheck().status()
                            : RawInputDateCrossSourceCheck.STATUS_NA);
            ui.setDailyReportOrderStatus(
                    dailyReportLookup.orderCompletionStatus(row.iraiNo()));
            ui.setPlanDayValues(row.planDayValues());
            ui.setSource(row);
            applyIssueState(ui, row, previousConfirmations);
            allRows.add(ui);
        }
        if (endIndex < sourceRows.size()) {
            Platform.runLater(
                    () ->
                            applyScanResultBatch(
                                    result,
                                    resetIssueConfirmations,
                                    previousConfirmations,
                                    endIndex));
            return;
        }
        finishApplyScanResult(result, resetIssueConfirmations);
    }

    private void finishApplyScanResult(ScanResult result, boolean resetIssueConfirmations) {
        lastScanWarnings =
                result.warnings().isEmpty() ? "" : String.join(" | ", result.warnings());
        applyFilter();
        updateStatusLabel();
        updateAladdinPlanSourceLabel();
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
        crossSourceRows.clear();
        planRows.clear();
        planContextRow = null;
        updateStage1GateLabel();
        notifyStage1GateChanged();
        if (resetIssueConfirmations && shell != null) {
            shell.appendLog("[pipeline-check] アラジン加工計画更新のため確認チェックをリセットしました。");
        }
        captureLastScannedAladdinPlanRevision();
        refreshAladdinPlanWatchState();
        flushPipelineCheckProgressUi();
        clearPipelineCheckProgress();
        applySourceExtensionOverlayFromWarnings(result.warnings());
        completeRefreshPreload(true, null);
    }

    private void applySourceExtensionOverlayFromWarnings(List<String> warnings) {
        Map<String, String> ui = shell != null ? shell.snapshotUiEnv() : Map.of();
        List<String> extErrs =
                new ArrayList<>(SourceFileExtensionPolicy.blockingMismatchMessages(ui));
        if (extErrs.isEmpty() && warnings != null) {
            for (String w : warnings) {
                if (w != null
                        && (w.contains("拡張子が不正") || w.contains("ファイル名が不正"))) {
                    extErrs.add(w);
                }
            }
        }
        if (extErrs.isEmpty()) {
            SourceExtensionErrorOverlay.clear(mainTableHost);
            return;
        }
        allRows.clear();
        applyFilter();
        updateStatusLabel();
        showSourceExtensionOverlay(extErrs, ui);
    }

    private void showSourceExtensionOverlay(List<String> extErrs, Map<String, String> ui) {
        List<Path> mismatchPaths = SourceFileExtensionPolicy.blockingMismatchPaths(ui);
        Window owner = shell != null ? shell.getPrimaryStage() : null;
        SourceExtensionErrorOverlay.show(
                mainTableHost,
                String.join("\n", extErrs),
                mismatchPaths,
                owner,
                deleted -> {
                    if (shell != null) {
                        for (Path p : deleted) {
                            shell.appendLog(
                                    "[pipeline-check] 不正拡張子ファイルを削除しました: " + p);
                        }
                        shell.refreshStage1PipelineCheckGate();
                    }
                    updateStage1GateLabel();
                });
    }

    /**
     * 他画面からアラジン加工計画ソースが再読込されたとき。
     * ファイル版が変わっていれば表示中の確認チェックをリセットする。
     */
    void onAladdinProcessingPlanSourceUpdated(Path sourceFile) {
        boolean planReloaded = registerAladdinPlanSourceReload(sourceFile);
        if (planReloaded) {
            resetAllIssueConfirmations();
            if (shell != null) {
                shell.appendLog("[pipeline-check] アラジン加工計画更新のため確認チェックをリセットしました。");
            }
        }
        Platform.runLater(this::refreshAladdinPlanWatchState);
    }

    private void resetAllIssueConfirmations() {
        for (MainRow row : allRows) {
            row.issueConfirmedProperty().set(false);
        }
        if (mainTable != null) {
            mainTable.refresh();
        }
        updateStage1GateLabel();
        notifyStage1GateChanged();
    }

    /**
     * アラジン加工計画ソースの版を登録する。
     *
     * @return 前回取り込みからファイルが更新されていれば {@code true}
     */
    boolean registerAladdinPlanSourceReload(Path sourceFile) {
        try {
            String revisionKey = aladdinPlanSourceRevisionKey(sourceFile);
            if (revisionKey.isBlank()) {
                return false;
            }
            boolean updated = !revisionKey.equals(lastAladdinPlanSourceRevisionKey);
            lastAladdinPlanSourceRevisionKey = revisionKey;
            return updated;
        } catch (IOException ex) {
            lastAladdinPlanSourceRevisionKey = "";
            return true;
        }
    }

    static String aladdinPlanSourceRevisionKey(Path sourceFile) throws IOException {
        if (sourceFile == null || !Files.isRegularFile(sourceFile)) {
            return "";
        }
        Path normalized = sourceFile.toAbsolutePath().normalize();
        return normalized + "|" + Files.getLastModifiedTime(normalized).toMillis();
    }

    static Optional<String> resolveCurrentAladdinPlanSourceRevisionKey(Map<String, String> ui) {
        try {
            Path dir = AppPaths.resolveTaskInputSourceDir(ui != null ? ui : Map.of());
            if (dir == null || !Files.isDirectory(dir)) {
                return Optional.empty();
            }
            Optional<Path> newest = NetworkSourceDirResolver.newestTaskInputFileInDirectory(dir, ui);
            if (newest.isEmpty()) {
                return Optional.empty();
            }
            Path file = newest.get();
            String low = file.getFileName().toString().toLowerCase(Locale.ROOT);
            if (low.endsWith(".pq") || low.endsWith(".parquet")) {
                return Optional.empty();
            }
            String key = aladdinPlanSourceRevisionKey(file);
            return key.isBlank() ? Optional.empty() : Optional.of(key);
        } catch (IOException ex) {
            return Optional.empty();
        }
    }

    boolean isAladdinPlanSourceNewerThanLastScan(Map<String, String> ui) {
        Optional<String> current = resolveCurrentAladdinPlanSourceRevisionKey(ui);
        if (current.isEmpty()) {
            return false;
        }
        if (lastScannedAladdinPlanSourceRevisionKey.isBlank()) {
            return scanApplied;
        }
        return !current.get().equals(lastScannedAladdinPlanSourceRevisionKey);
    }

    private void captureLastScannedAladdinPlanRevision() {
        if (shell == null) {
            return;
        }
        captureLastScannedAladdinPlanRevision(shell.snapshotUiEnv());
    }

    void captureLastScannedAladdinPlanRevisionForTest(Map<String, String> ui) {
        captureLastScannedAladdinPlanRevision(ui);
        scanApplied = true;
    }

    private void captureLastScannedAladdinPlanRevision(Map<String, String> ui) {
        resolveCurrentAladdinPlanSourceRevisionKey(ui)
                .ifPresent(key -> lastScannedAladdinPlanSourceRevisionKey = key);
    }

    private void startAladdinPlanWatch() {
        if (aladdinPlanWatchTimeline != null) {
            return;
        }
        aladdinPlanWatchTimeline =
                new Timeline(new KeyFrame(ALADDIN_PLAN_WATCH_INTERVAL, event -> tickAladdinPlanWatch()));
        aladdinPlanWatchTimeline.setCycleCount(Timeline.INDEFINITE);
        aladdinPlanWatchTimeline.play();
    }

    private void tickAladdinPlanWatch() {
        if (shell == null || refreshInProgress) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        String usedPlan = aladdinPlanSourcePath;
        String usedDaily = dailyReportLookup != null ? dailyReportLookup.sourcePath() : "";
        Thread worker =
                new Thread(
                        () -> {
                            boolean attention = refreshButtonNeedsAttention(ui, usedPlan, usedDaily);
                            Platform.runLater(
                                    () -> setRefreshButtonAttentionForAladdinPlanNewer(attention));
                        },
                        "pipeline-check-aladdin-watch");
        worker.setDaemon(true);
        worker.start();
    }

    private void refreshAladdinPlanWatchState() {
        if (shell == null || refreshInProgress) {
            return;
        }
        String usedDaily = dailyReportLookup != null ? dailyReportLookup.sourcePath() : "";
        boolean attention =
                refreshButtonNeedsAttention(shell.snapshotUiEnv(), aladdinPlanSourcePath, usedDaily);
        setRefreshButtonAttentionForAladdinPlanNewer(attention);
    }

    /**
     * 加工計画の版が走査時と異なる、または使用中の加工計画／加工日報が 取得データ最新表と異なるとき更新ボタンを点灯する。
     */
    boolean refreshButtonNeedsAttention(
            Map<String, String> ui, String usedPlanPath, String usedDailyPath) {
        return isAladdinPlanSourceNewerThanLastScan(ui)
                || RemoteDesktopLatestSourceFiles.pipelineCheckUsedPathsDifferFromLatest(
                        usedPlanPath, usedDailyPath, ui);
    }

    private void setRefreshButtonAttentionForAladdinPlanNewer(boolean newer) {
        if (refreshButton == null) {
            return;
        }
        if (refreshButtonGlow == null) {
            refreshButtonGlow = new ButtonAttentionGlow(refreshButton);
        }
        if (newer) {
            refreshButtonGlow.startIfIdle();
        } else {
            stopRefreshButtonGlow();
        }
    }

    private void stopRefreshButtonGlow() {
        ButtonAttentionGlow.stopAll(refreshButtonGlow);
    }

    private void applyIssueState(
            MainRow ui, PipelineStatusRow row, Map<String, Boolean> previousConfirmations) {
        List<IssueKind> issues = RequestFormPipelineIssueCheck.detect(row, aladdinJsonAvailable);
        ui.setIssueSummary(RequestFormPipelineIssueCheck.formatSummary(issues));
        ui.setHasIssues(!issues.isEmpty());
        boolean restoreConfirmed =
                previousConfirmations != null
                        && Boolean.TRUE.equals(previousConfirmations.get(ui.getIraiNo()));
        ui.issueConfirmedProperty().set(restoreConfirmed);
        ui.issueConfirmedProperty()
                .addListener(
                        (obs, oldVal, newVal) -> {
                            updateStage1GateLabel();
                            notifyStage1GateChanged();
                        });
    }

    /** 段階1実行可否（未走査・未確認・ソース拡張子不正があれば不可）。 */
    public Stage1GateStatus evaluateStage1Gate() {
        if (shell != null
                && !AppPaths.isRequestFormOriginalDirEnvConfigured(shell.snapshotUiEnv())) {
            return Stage1GateStatus.blocked(
                    "依頼書原本フォルダ（環境変数 "
                            + AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR
                            + "）を設定してから、原本転記・計画確認で「更新」を実行してください。",
                    "原本転記: 原本フォルダ未設定");
        }
        if (shell != null) {
            List<String> extErrs =
                    SourceFileExtensionPolicy.blockingMismatchMessages(shell.snapshotUiEnv());
            if (!extErrs.isEmpty()) {
                return Stage1GateStatus.blocked(
                        String.join("\n", extErrs), "ソース拡張子不正");
            }
        }
        if (!scanApplied) {
            return Stage1GateStatus.blocked(
                    "原本転記・計画確認で「更新」を実行し、問題の有無を確認してください。",
                    "原本転記: 未走査");
        }
        Stage1ConfirmationCounts counts = countStage1ConfirmationRequirements();
        if (counts.unconfirmedRequiringConfirmation() > 0) {
            return Stage1GateStatus.blocked(
                    "原本転記・計画確認: 要確認 "
                            + counts.requiringConfirmation()
                            + " 件のうち "
                            + counts.unconfirmedRequiringConfirmation()
                            + " 件が未確認です。該当行の「確認」にチェックを付けてください。"
                            + "（原反投入日不一致の3条件、またはアラジン計画なし＋納期/調整納期当日以降。"
                            + " 依頼NO先頭「2」の自社加工品は対象外）",
                    "原本転記: 未確認 "
                            + counts.unconfirmedRequiringConfirmation()
                            + "/"
                            + counts.requiringConfirmation()
                            + " 件");
        }
        return Stage1GateStatus.allowed();
    }

    private record Stage1ConfirmationCounts(
            int totalIssues, int requiringConfirmation, int unconfirmedRequiringConfirmation) {}

    private Stage1ConfirmationCounts countStage1ConfirmationRequirements() {
        int totalIssues = 0;
        int requiringConfirmation = 0;
        int unconfirmedRequiringConfirmation = 0;
        for (MainRow row : filteredRows) {
            if (!row.hasIssues()) {
                continue;
            }
            totalIssues++;
            if (!requiresStage1Confirmation(row)) {
                continue;
            }
            requiringConfirmation++;
            if (!row.isIssueConfirmed()) {
                unconfirmedRequiringConfirmation++;
            }
        }
        return new Stage1ConfirmationCounts(
                totalIssues, requiringConfirmation, unconfirmedRequiringConfirmation);
    }

    /** 調整納期が当日以降か。 */
    static boolean isAdjustDeliveryOnOrAfterTodayForRow(MainRow row) {
        return RequestFormPipelineStatusService.isAdjustDeliveryOnOrAfterToday(
                resolveAdjustDeliveryDate(row));
    }

    /** 日報列が未登録（―）。 */
    static boolean isDailyReportOrderStatusAbsent(MainRow row) {
        if (row == null) {
            return false;
        }
        String status = row.getDailyReportOrderStatus();
        return DAILY_REPORT_ORDER_STATUS_ABSENT.equals(status != null ? status.strip() : "");
    }

    /** 日報列が完了。 */
    static boolean isDailyReportOrderStatusComplete(MainRow row) {
        if (row == null) {
            return false;
        }
        String status = row.getDailyReportOrderStatus();
        return DAILY_REPORT_ORDER_STATUS_COMPLETE.equals(status != null ? status.strip() : "");
    }

    /** 原反投入日4ソース照合が不一致。 */
    static boolean isRawInputDateMismatchForRow(MainRow row) {
        if (row == null) {
            return false;
        }
        return RawInputDateCrossSourceCheck.STATUS_MISMATCH.equals(
                nullToEmpty(row.getRawInputDateMatchStatus()));
    }

    /** アラジン列が「なし」（JSON読込済みで当該依頼に計画行なし）。 */
    static boolean isAladdinPlanMissingForRow(MainRow row) {
        return row != null && "なし".equals(nullToEmpty(row.getAladdinStatus()));
    }

    /** 目次納期が当日以降か。 */
    static boolean isIndexDeliveryOnOrAfterTodayForRow(MainRow row) {
        LocalDate delivery = resolveIndexDeliveryDate(row);
        return delivery != null && !delivery.isBefore(LocalDate.now());
    }

    /** 目次納期の解釈（ソース行と表示列の両方から）。 */
    static LocalDate resolveIndexDeliveryDate(MainRow row) {
        if (row == null) {
            return null;
        }
        if (row.source() != null) {
            String fromSource = nullToEmpty(row.source().indexDeliveryDate());
            if (!fromSource.isBlank()) {
                LocalDate parsed = JuchuTransferValueNormalizer.parseLocalDate(fromSource.strip());
                if (parsed != null) {
                    return parsed;
                }
            }
        }
        String display = row.getIndexDeliveryDate();
        if (display == null || display.isBlank()) {
            return null;
        }
        return JuchuTransferValueNormalizer.parseLocalDate(display.strip());
    }

    /**
     * (A) 調整納期が当日以降・日報が―・原反投入日不一致の3条件すべて。
     */
    static boolean requiresStage1ConfirmationForRawInputMismatchTriplet(MainRow row) {
        return row != null
                && isAdjustDeliveryOnOrAfterTodayForRow(row)
                && isDailyReportOrderStatusAbsent(row)
                && isRawInputDateMismatchForRow(row);
    }

    /**
     * (B) アラジン計画なしかつ、原本納期または受注調整納期が当日以降。
     */
    static boolean requiresStage1ConfirmationForAladdinMissingWithDelivery(MainRow row) {
        if (row == null || !isAladdinPlanMissingForRow(row)) {
            return false;
        }
        PipelineStatusRow src = row.source();
        if (src == null) {
            return false;
        }
        boolean originalDeliveryPath =
                src.originalPresent() && isIndexDeliveryOnOrAfterTodayForRow(row);
        boolean juchuAdjustDeliveryPath =
                src.juchuRegistered() && isAdjustDeliveryOnOrAfterTodayForRow(row);
        return originalDeliveryPath || juchuAdjustDeliveryPath;
    }

    /**
     * 段階1前に確認チェックが必要な行（上記 (A) または (B)）。日報完了は対象外。
     */
    static boolean requiresStage1Confirmation(MainRow row) {
        if (row == null || isDailyReportOrderStatusComplete(row)) {
            return false;
        }
        if (RequestFormPipelineStatusService.isInHouseSelfProcessingIraiNo(row.getIraiNo())) {
            return false;
        }
        return requiresStage1ConfirmationForRawInputMismatchTriplet(row)
                || requiresStage1ConfirmationForAladdinMissingWithDelivery(row);
    }

    /**
     * 段階1要確認かつ未確認の行に確認チェックを付ける。新たに確認した件数を返す。
     */
    static int confirmAllRequiringConfirmation(Iterable<MainRow> rows) {
        if (rows == null) {
            return 0;
        }
        int count = 0;
        for (MainRow row : rows) {
            if (row != null && requiresStage1Confirmation(row) && !row.isIssueConfirmed()) {
                row.issueConfirmedProperty().set(true);
                count++;
            }
        }
        return count;
    }

    /** 調整納期の解釈（ソース行と表示列の両方から）。 */
    static LocalDate resolveAdjustDeliveryDate(MainRow row) {
        if (row == null) {
            return null;
        }
        if (row.source() != null) {
            LocalDate fromSource =
                    RequestFormPipelineStatusService.resolveAdjustDeliveryLocalDate(row.source());
            if (fromSource != null) {
                return fromSource;
            }
        }
        String display = row.getJuchuAdjustDeliveryDate();
        if (display == null || display.isBlank()) {
            return null;
        }
        return JuchuTransferValueNormalizer.parseLocalDate(display.strip());
    }

    public record Stage1GateStatus(boolean permitted, String message, String badgeMessage) {
        public static Stage1GateStatus allowed() {
            return new Stage1GateStatus(true, "", "");
        }

        public static Stage1GateStatus blocked(String message, String badgeMessage) {
            return new Stage1GateStatus(
                    false,
                    message != null ? message : "",
                    badgeMessage != null ? badgeMessage : "");
        }
    }

    private void notifyStage1GateChanged() {
        if (shell != null) {
            shell.refreshStage1PipelineCheckGate();
        }
    }

    private void updateStage1GateLabel() {
        if (stage1GateLabel != null) {
            if (!isRequestFormOriginalDirEnvConfigured()) {
                SourceExtensionErrorOverlay.clear(mainTableHost);
                stage1GateLabel.setText(
                        "段階1: 依頼書原本フォルダ未設定 — 環境変数タブで設定してください。");
                stage1GateLabel.getStyleClass().setAll("pipeline-check-stage1-gate-label", "warn");
            } else {
                List<String> extErrs =
                        shell != null
                                ? SourceFileExtensionPolicy.blockingMismatchMessages(
                                        shell.snapshotUiEnv())
                                : List.of();
                if (!extErrs.isEmpty()) {
                    stage1GateLabel.setText(
                            "段階1: ソース拡張子不正 — " + String.join(" / ", extErrs));
                    stage1GateLabel.getStyleClass().setAll("pipeline-check-stage1-gate-label", "warn");
                    showSourceExtensionOverlay(extErrs, shell.snapshotUiEnv());
                } else if (!scanApplied) {
                    SourceExtensionErrorOverlay.clear(mainTableHost);
                    stage1GateLabel.setText(
                            "段階1: 未走査 — 「更新」で照合してから実行してください。");
                    stage1GateLabel.getStyleClass().setAll("pipeline-check-stage1-gate-label", "warn");
                } else {
                    SourceExtensionErrorOverlay.clear(mainTableHost);
                    applyStage1GateLabelFromConfirmationCounts();
                }
            }
        }
        refreshConfirmAllIssuesButtonState();
    }

    private void applyStage1GateLabelFromConfirmationCounts() {
        Stage1ConfirmationCounts counts = countStage1ConfirmationRequirements();
        if (counts.totalIssues() == 0) {
            stage1GateLabel.setText("段階1: 問題なし — 実行できます。");
            stage1GateLabel.getStyleClass().setAll("pipeline-check-stage1-gate-label", "ok");
        } else if (counts.unconfirmedRequiringConfirmation() > 0) {
            stage1GateLabel.setText(
                    "段階1: 要確認 "
                            + counts.requiringConfirmation()
                            + " 件（未確認 "
                            + counts.unconfirmedRequiringConfirmation()
                            + " 件）— 確認チェック対象行のみ要確認");
            stage1GateLabel.getStyleClass().setAll("pipeline-check-stage1-gate-label", "warn");
        } else if (counts.requiringConfirmation() == 0) {
            stage1GateLabel.setText(
                    "段階1: 問題 "
                            + counts.totalIssues()
                            + " 件（確認チェック対象外あり）— 実行できます。");
            stage1GateLabel.getStyleClass().setAll("pipeline-check-stage1-gate-label", "ok");
        } else {
            stage1GateLabel.setText(
                    "段階1: 要確認 "
                            + counts.requiringConfirmation()
                            + " 件 — すべて確認済み。実行できます。");
            stage1GateLabel.getStyleClass().setAll("pipeline-check-stage1-gate-label", "ok");
        }
    }

    private void refreshConfirmAllIssuesButtonState() {
        if (confirmAllVisibleIssuesButton == null) {
            return;
        }
        if (!scanApplied || filteredRows == null) {
            confirmAllVisibleIssuesButton.setDisable(true);
            return;
        }
        confirmAllVisibleIssuesButton.setDisable(
                countStage1ConfirmationRequirements().unconfirmedRequiringConfirmation() <= 0);
    }

    private void showDetail(MainRow row) {
        mismatchRows.clear();
        crossSourceRows.clear();
        planRows.clear();
        planContextRow = row;
        if (row == null || row.source() == null) {
            return;
        }
        PipelineStatusRow src = row.source();
        populateCrossSourceRows(src.rawInputDateCrossCheck());
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
    private void onConfirmAllVisibleIssuesButtonAction() {
        if (!scanApplied || filteredRows == null) {
            return;
        }
        int confirmed = confirmAllRequiringConfirmation(filteredRows);
        if (confirmed <= 0) {
            return;
        }
        mainTable.refresh();
        if (shell != null) {
            shell.appendLog("[pipeline-check] 確認チェックを " + confirmed + " 件一括で付けました");
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

    @FXML
    private void onCopyCrossSourceTableButtonAction() {
        if (crossSourceRows.isEmpty()) {
            return;
        }
        CrossSourceCopyHeader header = CrossSourceCopyHeader.from(planContextRow);
        List<CrossSourceRow> rows = List.copyOf(crossSourceRows);
        String tsv = formatCrossSourceTableTsv(header, rows);
        ClipboardContent content = new ClipboardContent();
        content.putString(tsv);
        Clipboard.getSystemClipboard().setContent(content);
        if (shell != null) {
            shell.appendLog(
                    "[pipeline-check] 原反投入日4ソース照合 "
                            + rows.size()
                            + " 行をクリップボードへ表としてコピー"
                            + crossSourceHeaderLogSuffix(header));
        }
    }

    @FXML
    private void onCopyCrossSourceTableForEmailButtonAction() {
        if (crossSourceRows.isEmpty()) {
            return;
        }
        CrossSourceCopyHeader header = CrossSourceCopyHeader.from(planContextRow);
        List<CrossSourceRow> rows = List.copyOf(crossSourceRows);
        String tsv = formatCrossSourceTableTsv(header, rows);
        String html = formatCrossSourceTableHtml(header, rows);
        ClipboardTableSupport.copyTabularForRichTextPaste(tsv, html);
        if (shell != null) {
            shell.appendLog(
                    "[pipeline-check] 原反投入日4ソース照合 "
                            + rows.size()
                            + " 行をメール貼り付け用（HTML表）でクリップボードへコピー"
                            + crossSourceHeaderLogSuffix(header));
        }
    }

    private static String crossSourceHeaderLogSuffix(CrossSourceCopyHeader header) {
        if (header == null || header.iraiNo().isBlank()) {
            return "";
        }
        return "（依頼No: " + header.iraiNo() + "）";
    }

    /** 4ソース照合表クリップボード用ヘッダ（選択行）。 */
    record CrossSourceCopyHeader(String iraiNo, String matchStatus) {

        static CrossSourceCopyHeader from(MainRow row) {
            if (row == null) {
                return new CrossSourceCopyHeader("", "");
            }
            return new CrossSourceCopyHeader(
                    nullToEmpty(row.getIraiNo()), nullToEmpty(row.getRawInputDateMatchStatus()));
        }
    }

    static String formatCrossSourceTableTsv(
            CrossSourceCopyHeader header, List<CrossSourceRow> rows) {
        StringBuilder sb = new StringBuilder();
        appendCrossSourceCopyHeaderTsv(sb, header);
        if (!sb.isEmpty()) {
            sb.append('\n');
        }
        sb.append("ソース").append('\t').append("原反投入日").append('\t').append("照合");
        for (CrossSourceRow row : rows) {
            sb.append('\n');
            appendTsvCell(sb, row.getSource());
            sb.append('\t');
            appendTsvCell(sb, row.getValue());
            sb.append('\t');
            appendTsvCell(sb, row.getStatus());
        }
        return sb.toString();
    }

    static String formatCrossSourceTableHtml(
            CrossSourceCopyHeader header, List<CrossSourceRow> rows) {
        StringBuilder sb = new StringBuilder();
        appendCrossSourceCopyHeaderHtml(sb, header);
        sb.append(
                "<table border=\"1\" cellspacing=\"0\" cellpadding=\"4\""
                        + " style=\"border-collapse:collapse;font-family:'Meiryo UI',sans-serif;font-size:11pt;\">");
        sb.append("<thead><tr>");
        for (String columnTitle : List.of("ソース", "原反投入日", "照合")) {
            sb.append("<th style=\"background:#D9E1F2;padding:4px 8px;text-align:left;\">")
                    .append(ClipboardTableSupport.escapeHtml(columnTitle))
                    .append("</th>");
        }
        sb.append("</tr></thead><tbody>");
        for (CrossSourceRow row : rows) {
            sb.append("<tr>");
            appendHtmlCell(sb, row.getSource());
            appendHtmlCell(sb, row.getValue());
            appendHtmlCell(sb, row.getStatus());
            sb.append("</tr>");
        }
        sb.append("</tbody></table>");
        return sb.toString();
    }

    private static void appendCrossSourceCopyHeaderTsv(
            StringBuilder sb, CrossSourceCopyHeader header) {
        if (header == null) {
            return;
        }
        appendHeaderTsvLineRequired(sb, "依頼No", header.iraiNo());
        appendHeaderTsvLineRequired(sb, "投入日一致", header.matchStatus());
    }

    private static void appendCrossSourceCopyHeaderHtml(
            StringBuilder sb, CrossSourceCopyHeader header) {
        if (header == null) {
            return;
        }
        List<String[]> lines = new ArrayList<>();
        if (!header.iraiNo().isBlank()) {
            lines.add(new String[] {"依頼No", header.iraiNo()});
        }
        if (!header.matchStatus().isBlank()) {
            lines.add(new String[] {"投入日一致", header.matchStatus()});
        }
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
                    case "rawInputDateMatchStatus" -> row.getRawInputDateMatchStatus();
                    case "dailyReportOrderStatus" -> row.getDailyReportOrderStatus();
                    case "issueSummary" -> row.getIssueSummary();
                    case "issueConfirmedDisplay" -> row.getIssueConfirmedDisplay();
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
        boolean showRawInputMismatchOnly =
                showRawInputMismatchOnlyCheck != null
                        && showRawInputMismatchOnlyCheck.isSelected();

        filteredRows.setPredicate(
                row -> {
                    PipelineStatusRow src = row.source();
                    if (src == null) {
                        return false;
                    }
                    if (requiresStage1Confirmation(row)) {
                        return matchesPipelineCheckQuickSearch(row, src, q);
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
                                            RequestFormPipelineStatusService
                                                    .resolveAdjustDeliveryLocalDate(src))) {
                        return false;
                    }
                    if (showDailyReportIncompleteOnly
                            && "完了".equals(nullToEmpty(row.getDailyReportOrderStatus()))) {
                        return false;
                    }
                    if (showRawInputMismatchOnly
                            && !RawInputDateCrossSourceCheck.STATUS_MISMATCH.equals(
                                    nullToEmpty(row.getRawInputDateMatchStatus()))) {
                        return false;
                    }
                    return matchesPipelineCheckQuickSearch(row, src, q);
                });
        updateStatusLabel();
        refreshConfirmAllIssuesButtonState();
    }

    /** 依頼No・ユーザー名のクイック検索（空なら常に一致）。 */
    private static boolean matchesPipelineCheckQuickSearch(
            MainRow row, PipelineStatusRow src, String q) {
        if (q == null || q.isEmpty()) {
            return true;
        }
        String irai =
                row.getIraiNo() != null ? row.getIraiNo().toLowerCase(Locale.ROOT) : "";
        String user = src.user() != null ? src.user().toLowerCase(Locale.ROOT) : "";
        return irai.contains(q) || user.contains(q);
    }

    private int resolveJuchuInputHideDays() {
        if (juchuInputHideDaysSpinner == null || juchuInputHideDaysSpinner.getValue() == null) {
            return RequestFormPipelineStatusService.DEFAULT_JUCHU_INPUT_DATE_HIDE_DAYS;
        }
        return Math.max(0, juchuInputHideDaysSpinner.getValue());
    }

    /** 走査時に古い受注行をスキップする日数（チェック OFF なら 0）。 */
    private int resolveJuchuInputHideDaysForScan() {
        if (hideOldJuchuInputCheck == null || !hideOldJuchuInputCheck.isSelected()) {
            return 0;
        }
        return resolveJuchuInputHideDays();
    }

    /**
     * 受注入力日フィルタは走査時にも適用する。設定変更後は再走査して結果を揃える。
     */
    private void onJuchuInputHideScanSettingChanged() {
        applyFilter();
        if (scanApplied && !refreshInProgress) {
            startRefresh(true);
        }
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
                            RequestFormPipelineStatusService.resolveAdjustDeliveryLocalDate(
                                    src))) {
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

    private void updateAladdinPlanSourceLabel() {
        if (aladdinPlanSourceLabel == null) {
            return;
        }
        boolean hasPath = aladdinPlanSourcePath != null && !aladdinPlanSourcePath.isBlank();
        if (!hasPath) {
            aladdinPlanSourceLabel.setText("加工計画: （未読込）");
        } else {
            aladdinPlanSourceLabel.setText("加工計画: " + aladdinPlanSourcePath);
        }
        aladdinPlanSourceLabel.setManaged(true);
        aladdinPlanSourceLabel.setVisible(true);
        if (openAladdinPlanSourceFolderButton != null) {
            openAladdinPlanSourceFolderButton.setDisable(!hasPath);
        }
    }

    private void updateDailyReportSourceLabel() {
        if (dailyReportSourceLabel == null) {
            return;
        }
        String path = dailyReportLookup != null ? dailyReportLookup.sourcePath() : "";
        boolean hasPath = !path.isBlank();
        if (!hasPath) {
            dailyReportSourceLabel.setText("加工日報: （未読込）");
            dailyReportSourceLabel.setManaged(true);
            dailyReportSourceLabel.setVisible(true);
            if (openDailyReportSourceFolderButton != null) {
                openDailyReportSourceFolderButton.setDisable(true);
            }
            return;
        }
        dailyReportSourceLabel.setText("加工日報: " + path);
        dailyReportSourceLabel.setManaged(true);
        dailyReportSourceLabel.setVisible(true);
        if (openDailyReportSourceFolderButton != null) {
            openDailyReportSourceFolderButton.setDisable(false);
        }
    }

    @FXML
    private void onOpenAladdinPlanSourceFolderAction() {
        openSourceFileParentFolder(aladdinPlanSourcePath, "加工計画");
    }

    @FXML
    private void onOpenDailyReportSourceFolderAction() {
        String path = dailyReportLookup != null ? dailyReportLookup.sourcePath() : "";
        openSourceFileParentFolder(path, "加工日報");
    }

    /** 表示中ソース CSV の親フォルダをエクスプローラー等で開く。 */
    private void openSourceFileParentFolder(String rawPath, String label) {
        if (rawPath == null || rawPath.isBlank()) {
            warnOpenSourceFolder(label + "のパスが未設定です。");
            return;
        }
        Path filePath;
        try {
            filePath = Path.of(rawPath.trim());
        } catch (Exception ex) {
            warnOpenSourceFolder(label + "のパスが無効です: " + rawPath);
            return;
        }
        Path dir = resolveOpenableDirectory(filePath);
        if (dir == null) {
            warnOpenSourceFolder(label + "のフォルダを開けません: " + rawPath);
            return;
        }
        try {
            openDirectoryInOs(dir);
            if (shell != null) {
                shell.appendLog("[pipeline-check] " + label + "フォルダを開きました: " + dir);
            }
        } catch (IOException ex) {
            String msg =
                    ex.getMessage() != null && !ex.getMessage().isBlank()
                            ? ex.getMessage()
                            : ex.toString();
            warnOpenSourceFolder(label + "のフォルダを開けませんでした: " + msg);
        }
    }

    private static Path resolveOpenableDirectory(Path path) {
        if (path == null) {
            return null;
        }
        try {
            if (Files.isDirectory(path)) {
                return path.toAbsolutePath().normalize();
            }
            Path parent = path.getParent();
            if (parent != null && Files.isDirectory(parent)) {
                return parent.toAbsolutePath().normalize();
            }
            if (Files.isRegularFile(path)) {
                Path p = path.getParent();
                return p != null ? p.toAbsolutePath().normalize() : null;
            }
        } catch (Exception ignored) {
            // UNC 等で存在確認が失敗しても親を試す
        }
        Path parent = path.getParent();
        return parent;
    }

    private static void openDirectoryInOs(Path dir) throws IOException {
        if (dir == null) {
            throw new IOException("directory is null");
        }
        String os = System.getProperty("os.name", "").toLowerCase(Locale.ROOT);
        if (os.contains("windows")) {
            new ProcessBuilder("explorer.exe", dir.toString()).start();
            return;
        }
        if (!Desktop.isDesktopSupported()
                || !Desktop.getDesktop().isSupported(Desktop.Action.OPEN)) {
            throw new IOException("Desktop OPEN is not supported");
        }
        Desktop.getDesktop().open(dir.toFile());
    }

    private void warnOpenSourceFolder(String message) {
        Alert alert = new Alert(Alert.AlertType.WARNING, message);
        alert.setTitle("フォルダを開く");
        alert.setHeaderText(null);
        Window owner = shell != null ? shell.getPrimaryStage() : null;
        if (owner != null) {
            alert.initOwner(owner);
        }
        alert.showAndWait();
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
            if ("rawInputDateMatchStatus".equals(def.property())) {
                installRawInputMatchCellFactory(column);
            } else if ("issueSummary".equals(def.property())) {
                installIssueSummaryCellFactory(column);
            } else if ("issueConfirmedDisplay".equals(def.property())) {
                installIssueConfirmCellFactory(column);
            }
            columns.add(column);
        }
        suppressColumnPersistence.set(true);
        try {
            mainTable.getColumns().setAll(columns);
            mainTable.setColumnResizePolicy(TableView.UNCONSTRAINED_RESIZE_POLICY);
            installMainTableRowFactory();
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

    private void populateCrossSourceRows(RawInputDateCrossSourceCheck.CrossSourceResult result) {
        if (result == null) {
            return;
        }
        String status = result.status();
        RawInputDateCrossSourceCheck.SourceValues v = result.values();
        crossSourceRows.add(
                new CrossSourceRow(
                        RawInputDateCrossSourceCheck.Source.ALADDIN.label(), v.aladdin(), status));
        crossSourceRows.add(
                new CrossSourceRow(
                        RawInputDateCrossSourceCheck.Source.JUCHU.label(), v.juchu(), status));
        crossSourceRows.add(
                new CrossSourceRow(
                        RawInputDateCrossSourceCheck.Source.INDEX.label(), v.index(), status));
        crossSourceRows.add(
                new CrossSourceRow(
                        RawInputDateCrossSourceCheck.Source.SHEET.label(), v.sheet(), status));
    }

    private void setupCrossSourceColumns() {
        crossSourceTable.getColumns().setAll(
                colCrossSource("ソース", "source", 150),
                colCrossSource("原反投入日", "value", 160),
                colCrossSource("照合", "status", 72));
    }

    private static void clearPipelineCheckCellAccent(javafx.scene.control.TableCell<?, ?> cell) {
        cell.getStyleClass()
                .removeAll(
                        "pipeline-check-cell-mismatch",
                        "pipeline-check-cell-match",
                        "pipeline-check-cell-issue");
        cell.setStyle("");
    }

    private static void applyRawInputDateMatchStatusCellStyle(
            javafx.scene.control.TableCell<?, ?> cell, String status) {
        clearPipelineCheckCellAccent(cell);
        if (RawInputDateCrossSourceCheck.STATUS_MISMATCH.equals(status)) {
            cell.getStyleClass().add("pipeline-check-cell-mismatch");
        } else if (RawInputDateCrossSourceCheck.STATUS_MATCH.equals(status)) {
            cell.getStyleClass().add("pipeline-check-cell-match");
        }
    }

    /** 下段4ソース照合用（明背景テーマのままインライン色を使用）。 */
    private static void applyCrossSourceStatusCellStyle(
            javafx.scene.control.TableCell<?, ?> cell, String status) {
        if (RawInputDateCrossSourceCheck.STATUS_MISMATCH.equals(status)) {
            cell.setStyle(
                    "-fx-background-color: #FCE4E4; -fx-control-inner-background: #FCE4E4;"
                            + " -fx-text-fill: #C00000;");
        } else if (RawInputDateCrossSourceCheck.STATUS_MATCH.equals(status)) {
            cell.setStyle(
                    "-fx-background-color: #E2EFDA; -fx-control-inner-background: #E2EFDA;"
                            + " -fx-text-fill: #375623;");
        } else {
            cell.setStyle("");
        }
    }

    private static void installRawInputMatchCellFactory(TableColumn<MainRow, String> column) {
        column.setCellFactory(
                col ->
                        new javafx.scene.control.TableCell<>() {
                            @Override
                            protected void updateItem(String item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty || item == null || item.isBlank()) {
                                    setText("");
                                    clearPipelineCheckCellAccent(this);
                                    return;
                                }
                                setText(item);
                                applyRawInputDateMatchStatusCellStyle(this, item);
                            }
                        });
    }

    private static void installIssueSummaryCellFactory(TableColumn<MainRow, String> column) {
        column.setCellFactory(
                col ->
                        new javafx.scene.control.TableCell<>() {
                            @Override
                            protected void updateItem(String item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty) {
                                    setText(null);
                                    clearPipelineCheckCellAccent(this);
                                    return;
                                }
                                MainRow row = getTableRow() != null ? getTableRow().getItem() : null;
                                setText(item != null ? item : "");
                                clearPipelineCheckCellAccent(this);
                                if (row != null && row.hasIssues()) {
                                    if (requiresStage1Confirmation(row)) {
                                        getStyleClass().add("pipeline-check-cell-issue");
                                    } else {
                                        getStyleClass().add("pipeline-check-cell-exempt");
                                    }
                                }
                            }
                        });
    }

    private void installIssueConfirmCellFactory(TableColumn<MainRow, String> column) {
        column.setCellFactory(
                col ->
                        new javafx.scene.control.TableCell<>() {
                            private final CheckBox checkBox = new CheckBox();

                            {
                                checkBox.setOnAction(
                                        event -> {
                                            MainRow row =
                                                    getTableRow() != null
                                                            ? getTableRow().getItem()
                                                            : null;
                                            if (row != null && requiresStage1Confirmation(row)) {
                                                row.issueConfirmedProperty().set(checkBox.isSelected());
                                            }
                                        });
                            }

                            @Override
                            protected void updateItem(String item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty) {
                                    setGraphic(null);
                                    setText(null);
                                    return;
                                }
                                MainRow row = getTableRow() != null ? getTableRow().getItem() : null;
                                if (!requiresStage1Confirmation(row)) {
                                    setGraphic(null);
                                    setText("―");
                                    return;
                                }
                                checkBox.setSelected(row.isIssueConfirmed());
                                setGraphic(checkBox);
                                setText(null);
                            }
                        });
    }

    private void installMainTableRowFactory() {
        mainTable.setRowFactory(
                tv ->
                        new javafx.scene.control.TableRow<>() {
                            @Override
                            protected void updateItem(MainRow item, boolean empty) {
                                super.updateItem(item, empty);
                                getStyleClass()
                                        .removeAll(
                                                "pipeline-check-row-unconfirmed",
                                                "pipeline-check-row-confirmation-exempt");
                                if (empty || item == null) {
                                    return;
                                }
                                if (requiresStage1Confirmation(item)) {
                                    if (!item.isIssueConfirmed()) {
                                        getStyleClass().add("pipeline-check-row-unconfirmed");
                                    }
                                } else if (item.hasIssues()) {
                                    getStyleClass().add("pipeline-check-row-confirmation-exempt");
                                }
                            }
                        });
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

    private static TableColumn<CrossSourceRow, String> colCrossSource(
            String title, String prop, double width) {
        TableColumn<CrossSourceRow, String> c = new TableColumn<>(title);
        c.setCellValueFactory(new PropertyValueFactory<>(prop));
        c.setPrefWidth(width);
        c.setCellFactory(
                col ->
                        new javafx.scene.control.TableCell<>() {
                            @Override
                            protected void updateItem(String item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty) {
                                    setText(null);
                                    setStyle("");
                                    return;
                                }
                                setText(item);
                                CrossSourceRow row =
                                        getTableRow() != null ? getTableRow().getItem() : null;
                                applyCrossSourceStatusCellStyle(
                                        this, row != null ? row.getStatus() : "");
                            }
                        });
        return c;
    }

    private static TableColumn<PlanRow, String> colPlan(String title, String prop, double width) {
        TableColumn<PlanRow, String> c = new TableColumn<>(title);
        c.setCellValueFactory(new PropertyValueFactory<>(prop));
        c.setPrefWidth(width);
        return c;
    }
}
