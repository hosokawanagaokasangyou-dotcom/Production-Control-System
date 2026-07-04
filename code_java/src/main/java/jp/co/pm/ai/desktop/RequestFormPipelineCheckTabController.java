package jp.co.pm.ai.desktop;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import javafx.application.Platform;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.collections.transformation.FilteredList;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.Label;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;

import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup;
import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup.PlanEntry;
import jp.co.pm.ai.desktop.reconciliation.JuchuTransferCoverageCheck.ColumnCheck;
import jp.co.pm.ai.desktop.reconciliation.JuchuHeaderAliasRegistry;
import jp.co.pm.ai.desktop.reconciliation.RequestFormPipelineStatusService;
import jp.co.pm.ai.desktop.reconciliation.RequestFormPipelineStatusService.PipelineStatusRow;
import jp.co.pm.ai.desktop.reconciliation.RequestFormPipelineStatusService.ScanResult;

/** 依頼書原本の受注転記率・アラジン加工計画の有無を一覧表示する。 */
public final class RequestFormPipelineCheckTabController {

    private static final int PLAN_DAY_COLUMNS =
            AladdinShapedPlanQtyLookup.PIPELINE_CHECK_PLAN_DAY_COLUMNS;

    private static final String HINT_TEXT =
            "依頼書原本フォルダ内の Excel 原本を走査し、受注ファイルへの転記状況と"
                    + " shaped_aladdin_plan.json 上のアラジン加工計画を照合します。"
                    + " 転記率は原本に値がある転記対象列を分母とします。"
                    + " アラジン計画は先頭7日分を①〜⑦列に表示（セル例: 7/3 100m）。"
                    + " 受注入力日フィルタは既定30日（変更可）。";

    public static final class MainRow {
        private String iraiNo;
        private String originalFile;
        private String juchuStatus;
        private String rateDisplay;
        private String mismatchCount;
        private String contractNoStatus;
        private String aladdinStatus;
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

        public String getJuchuStatus() {
            return juchuStatus;
        }

        public void setJuchuStatus(String juchuStatus) {
            this.juchuStatus = juchuStatus;
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

        PlanRow(PlanEntry entry) {
            machineName.set(entry.machineName());
            processName.set(entry.processName());
            dateYmd.set(entry.dateYmd());
            planMeters.set(formatMeters(entry.planMeters()));
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

        private static String formatMeters(double m) {
            if (Math.abs(m - Math.rint(m)) < 1e-9) {
                return String.valueOf((long) Math.rint(m));
            }
            return String.valueOf(m);
        }
    }

    @FXML
    private Button refreshButton;

    @FXML
    private TextField filterField;

    @FXML
    private CheckBox hideNoOriginalCheck;

    @FXML
    private CheckBox hideOldJuchuInputCheck;

    @FXML
    private Spinner<Integer> juchuInputHideDaysSpinner;

    @FXML
    private Label statusLabel;

    @FXML
    private Label hintLabel;

    @FXML
    private TableView<MainRow> mainTable;

    @FXML
    private TableView<MismatchRow> mismatchTable;

    @FXML
    private TableView<PlanRow> planTable;

    private MainShellController shell;

    private final ObservableList<MainRow> allRows = FXCollections.observableArrayList();
    private FilteredList<MainRow> filteredRows;
    private final ObservableList<MismatchRow> mismatchRows = FXCollections.observableArrayList();
    private final ObservableList<PlanRow> planRows = FXCollections.observableArrayList();

    private boolean aladdinJsonAvailable = true;
    private List<String> planDateHeaders = List.of();
    private String lastScanWarnings = "";

    @FXML
    private void initialize() {
        hintLabel.setText(HINT_TEXT);
        filteredRows = new FilteredList<>(allRows, row -> true);
        mainTable.setItems(filteredRows);
        mismatchTable.setItems(mismatchRows);
        planTable.setItems(planRows);
        VBox.setVgrow(mainTable, Priority.ALWAYS);

        setupMainColumns(List.of());
        setupMismatchColumns();
        setupPlanColumns();

        filterField.textProperty().addListener((obs, oldVal, newVal) -> applyFilter());
        hideNoOriginalCheck.selectedProperty().addListener((obs, o, n) -> applyFilter());
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
    }

    @FXML
    private void onRefreshButtonAction() {
        if (shell == null) {
            statusLabel.setText("シェル未接続");
            return;
        }
        refreshButton.setDisable(true);
        statusLabel.setText("走査中...");
        JuchuHeaderAliasRegistry registry = shell.snapshotJuchuHeaderAliasRegistryForExport();
        Thread worker =
                new Thread(
                        () -> {
                            ScanResult result =
                                    RequestFormPipelineStatusService.scan(
                                            shell.snapshotUiEnv(), registry);
                            Platform.runLater(
                                    () -> {
                                        refreshButton.setDisable(false);
                                        applyScanResult(result);
                                    });
                        },
                        "request-form-pipeline-check");
        worker.setDaemon(true);
        worker.start();
    }

    private void applyScanResult(ScanResult result) {
        allRows.clear();
        aladdinJsonAvailable = result.aladdinJsonAvailable();
        planDateHeaders = result.planDateHeaders() != null ? result.planDateHeaders() : List.of();
        setupMainColumns(planDateHeaders);

        for (PipelineStatusRow row : result.rows()) {
            MainRow ui = new MainRow();
            ui.setIraiNo(row.iraiNo());
            ui.setOriginalFile(
                    row.originalPresent()
                            ? row.originalFileName()
                            : "（依頼書原本なし）");
            ui.setJuchuStatus(row.juchuRegistered() ? "あり" : "なし");
            ui.setRateDisplay(row.rateDisplay());
            ui.setMismatchCount(String.valueOf(row.mismatchCount()));
            ui.setContractNoStatus(row.contractNoStatus());
            if (!aladdinJsonAvailable) {
                ui.setAladdinStatus("未確認");
            } else {
                ui.setAladdinStatus(row.aladdinPresent() ? "あり" : "なし");
            }
            ui.setPlanDayValues(row.planDayValues());
            ui.setSource(row);
            allRows.add(ui);
        }
        lastScanWarnings =
                result.warnings().isEmpty() ? "" : String.join(" | ", result.warnings());
        applyFilter();
        updateStatusLabel();

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
    }

    private void showDetail(MainRow row) {
        mismatchRows.clear();
        planRows.clear();
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
            planRows.add(new PlanRow(entry));
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
        if (!planDateHeaders.isEmpty()) {
            status.append(" | 計画日列: ");
            for (int i = 0; i < planDateHeaders.size(); i++) {
                if (i > 0) {
                    status.append(", ");
                }
                status.append(AladdinShapedPlanQtyLookup.circledSlotColumnLabel(i))
                        .append('=')
                        .append(
                                AladdinShapedPlanQtyLookup.shortPlanDateColumnLabel(
                                        planDateHeaders.get(i)));
            }
        }
        if (!lastScanWarnings.isEmpty()) {
            status.append(" | ").append(lastScanWarnings);
        }
        statusLabel.setText(status.toString());
    }

    private void setupMainColumns(List<String> dateHeaders) {
        List<TableColumn<MainRow, String>> cols = new ArrayList<>();
        cols.add(col("依頼No", "iraiNo", 90));
        cols.add(col("原本", "originalFile", 140));
        cols.add(col("受注", "juchuStatus", 52));
        cols.add(col("転記率", "rateDisplay", 100));
        cols.add(col("未一致", "mismatchCount", 52));
        cols.add(col("契約NO", "contractNoStatus", 88));
        cols.add(col("Aladdin", "aladdinStatus", 64));
        for (int i = 0; i < PLAN_DAY_COLUMNS; i++) {
            String title = AladdinShapedPlanQtyLookup.circledSlotColumnLabel(i);
            cols.add(col(title, "planDay" + i, 88));
        }
        mainTable.getColumns().setAll(cols);
    }

    private void setupMismatchColumns() {
        mismatchTable.getColumns().setAll(
                colMismatch("フォーム項目", "formLabel", 220),
                colMismatch("原本値", "originalValue", 160),
                colMismatch("受注値", "juchuValue", 160));
    }

    private void setupPlanColumns() {
        planTable.getColumns().setAll(
                colPlan("機械名", "machineName", 140),
                colPlan("工程名", "processName", 100),
                colPlan("日付", "dateYmd", 100),
                colPlan("計画m", "planMeters", 80));
    }

    private static TableColumn<MainRow, String> col(String title, String prop, double width) {
        TableColumn<MainRow, String> c = new TableColumn<>(title);
        c.setCellValueFactory(new PropertyValueFactory<>(prop));
        c.setPrefWidth(width);
        c.setMinWidth(width * 0.6);
        return c;
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
