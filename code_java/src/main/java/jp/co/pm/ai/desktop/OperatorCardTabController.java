package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.Parent;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ComboBox;
import jp.co.pm.ai.desktop.ui.InlineMonthCalendarPane;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.TextField;
import javafx.scene.control.TextFormatter;
import javafx.scene.layout.StackPane;
import javafx.scene.text.Font;
import javafx.print.PageLayout;
import javafx.print.PageOrientation;
import javafx.print.Paper;
import javafx.print.Printer;
import javafx.print.PrinterJob;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.Stage2OutputNaming;
import jp.co.pm.ai.desktop.io.JsonTableIo;
import jp.co.pm.ai.desktop.io.JsonTableIo.SheetTable;
import jp.co.pm.ai.desktop.print.OperatorCardDocumentBuilder;
import jp.co.pm.ai.desktop.print.OperatorCardDocumentBuilder.OperatorCardBuildException;
import jp.co.pm.ai.desktop.print.OperatorCardPage;
import jp.co.pm.ai.desktop.print.OperatorCardPreviewFactory;

/**
 * Operator dispatch card (A4 preview / print) tab — layout {@code OperatorCardTab.fxml}.
 */
public final class OperatorCardTabController {

    @FXML
    private Button syncLatestButton;

    @FXML
    private Button refreshPreviewButton;

    @FXML
    private Button printButton;

    @FXML
    private TextField memberJsonField;

    @FXML
    private TextField dispatchJsonField;

    @FXML
    private Button browseMemberButton;

    @FXML
    private Button browseDispatchButton;

    @FXML
    private InlineMonthCalendarPane startDateCalendar;

    @FXML
    private Spinner<Integer> dayCountSpinner;

    @FXML
    private ComboBox<String> operatorCombo;

    @FXML
    private ComboBox<String> fontCombo;

    @FXML
    private CheckBox printAllOperatorsCheckBox;

    @FXML
    private Label statusLabel;

    @FXML
    private StackPane previewHost;

    private MainShellController shell;

    private Stage ownerStage;

    private Map<String, SheetTable> cachedMemberSheets = Map.of();

    @FXML
    private void initialize() {
        installDayCountSpinner();
        if (startDateCalendar != null) {
            startDateCalendar.setSelectedDate(LocalDate.now());
            startDateCalendar
                    .selectedDateProperty()
                    .addListener(
                            (obs, previousDate, newDate) -> {
                                if (!Objects.equals(previousDate, newDate)) {
                                    rebuildPreview();
                                }
                            });
        }
        if (dayCountSpinner != null) {
            dayCountSpinner
                    .valueProperty()
                    .addListener(
                            (obs, previousDays, newDays) -> {
                                if (!Objects.equals(previousDays, newDays)) {
                                    rebuildPreview();
                                }
                            });
        }
        if (operatorCombo != null) {
            operatorCombo
                    .valueProperty()
                    .addListener(
                            (obs, previousOp, newOp) -> {
                                if (!Objects.equals(previousOp, newOp)) {
                                    rebuildPreview();
                                }
                            });
        }
        if (previewHost != null) {
            previewHost.setAlignment(Pos.TOP_CENTER);
            Label placeholder =
                    new Label(
                            "member_schedule*.json と結果_配台表.json"
                                    + " を指定し、プレビュー更新を押してください。");
            previewHost.getChildren().setAll(placeholder);
        }
    }

    private void installDayCountSpinner() {
        if (dayCountSpinner == null) {
            return;
        }
        int min = OperatorCardDocumentBuilder.MIN_DAY_COUNT;
        int max = OperatorCardDocumentBuilder.MAX_DAY_COUNT;
        int initial = OperatorCardDocumentBuilder.DEFAULT_DAY_COUNT;
        SpinnerValueFactory.IntegerSpinnerValueFactory vf =
                new SpinnerValueFactory.IntegerSpinnerValueFactory(min, max, initial);
        vf.setAmountToStepBy(1);
        dayCountSpinner.setValueFactory(vf);
        dayCountSpinner.setEditable(true);
        TextFormatter<Integer> formatter =
                new TextFormatter<>(
                        vf.getConverter(),
                        vf.getValue(),
                        change -> {
                            String text = change.getControlNewText();
                            if (text.isEmpty()) {
                                return change;
                            }
                            return text.matches("\\d{0,2}") ? change : null;
                        });
        dayCountSpinner.getEditor().setTextFormatter(formatter);
        dayCountSpinner
                .getEditor()
                .focusedProperty()
                .addListener(
                        (obs, wasFocused, isFocused) -> {
                            if (Boolean.FALSE.equals(isFocused)) {
                                commitDayCountSpinnerValue();
                            }
                        });
    }

    private void commitDayCountSpinnerValue() {
        if (dayCountSpinner == null || !dayCountSpinner.isEditable()) {
            return;
        }
        try {
            dayCountSpinner.commitValue();
        } catch (IllegalArgumentException ex) {
            Integer cur = dayCountSpinner.getValue();
            if (cur != null) {
                dayCountSpinner.getEditor().setText(Integer.toString(cur));
            }
        }
    }

    private int selectedDayCount() {
        commitDayCountSpinnerValue();
        if (dayCountSpinner == null) {
            return OperatorCardDocumentBuilder.DEFAULT_DAY_COUNT;
        }
        String raw = dayCountSpinner.getEditor().getText();
        if (raw != null) {
            String t = raw.trim();
            if (!t.isEmpty()) {
                try {
                    return OperatorCardDocumentBuilder.clampDayCount(Integer.parseInt(t));
                } catch (NumberFormatException ignored) {
                    // fall through
                }
            }
        }
        Integer val = dayCountSpinner.getValue();
        return val != null
                ? OperatorCardDocumentBuilder.clampDayCount(val)
                : OperatorCardDocumentBuilder.DEFAULT_DAY_COUNT;
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        this.ownerStage = shell.getPrimaryStage();
        populateFontCombo();
        Platform.runLater(this::applyDefaultPathsFromEnv);
    }

    private void applyDefaultPathsFromEnv() {
        if (shell == null || dispatchJsonField == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        Path disp = AppPaths.resolveResultDispatchTableJsonPath(ui);
        dispatchJsonField.setPromptText(disp.toString());
        if (dispatchJsonField.getText() == null || dispatchJsonField.getText().isBlank()) {
            if (Files.isRegularFile(disp)) {
                dispatchJsonField.setText(disp.toString());
            }
        }
    }

    private void populateFontCombo() {
        if (fontCombo == null) {
            return;
        }
        ObservableList<String> fam = FXCollections.observableArrayList(Font.getFamilies());
        FXCollections.sort(fam);
        fontCombo.setItems(fam);
        fontCombo.setValue(pickDefaultFont(fam));
    }

    static String pickDefaultFont(ObservableList<String> families) {
        if (families == null || families.isEmpty()) {
            return "SansSerif";
        }
        List<String> prefer =
                List.of(
                        "BIZ UDゴシック",
                        "BIZ UD Gothic",
                        "BIZ UDPゴシック",
                        "BIZ UDPGothic");
        for (String p : prefer) {
            if (families.contains(p)) {
                return p;
            }
        }
        if (families.contains("Meiryo UI")) {
            return "Meiryo UI";
        }
        return families.get(0);
    }

    @FXML
    private void onBrowseMemberJsonAction() {
        browseJson(memberJsonField);
    }

    @FXML
    private void onBrowseDispatchJsonAction() {
        browseJson(dispatchJsonField);
    }

    private void browseJson(TextField target) {
        FileChooser ch = new FileChooser();
        ch.setTitle("JSON");
        ch.getExtensionFilters().add(new FileChooser.ExtensionFilter("JSON", "*.json"));
        ch.getExtensionFilters().add(new FileChooser.ExtensionFilter("All", "*.*"));
        if (shell != null) {
            try {
                Map<String, String> ui = shell.snapshotUiEnv();
                Path dir = AppPaths.defaultPlanningOutputDir(ui);
                if (Files.isDirectory(dir)) {
                    ch.setInitialDirectory(dir.toFile());
                }
            } catch (Exception ignored) {
                // ignore
            }
        }
        java.io.File picked = ch.showOpenDialog(ownerStage);
        if (picked != null) {
            target.setText(picked.getAbsolutePath());
            reloadMemberCachesAndOperators();
        }
    }

    @FXML
    private void onSyncLatestButtonAction() {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        Path dir = AppPaths.defaultPlanningOutputDir(ui);
        try {
            Path mem = Stage2OutputNaming.newestPrimaryMemberJson(dir);
            Path dispDirFile = AppPaths.resolveResultDispatchTableJsonPath(ui);
            if (mem != null) {
                memberJsonField.setText(mem.toString());
            }
            if (Files.isRegularFile(dispDirFile)) {
                dispatchJsonField.setText(dispDirFile.toString());
            }
            if (mem == null && !Files.isRegularFile(dispDirFile)) {
                statusLabel.setText(
                        "最新 JSON が見つかりません: " + dir);
                return;
            }
            statusLabel.setText(
                    "sync: member="
                            + (mem != null ? mem.getFileName() : "-")
                            + ", dispatch="
                            + (Files.isRegularFile(dispDirFile) ? dispDirFile.getFileName() : "-"));
        } catch (Exception ex) {
            statusLabel.setText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
            return;
        }
        reloadMemberCachesAndOperators();
    }

    @FXML
    private void onRefreshPreviewButtonAction() {
        rebuildPreview();
    }

    @FXML
    private void onPrintButtonAction() {
        printCards();
    }

    /**
     * 段階2／段階3 完了後: 正規 {@code 結果_配台表.json} を反映し、member キャッシュを再読込してプレビューを更新する。
     */
    void syncAfterPipelineArtifactRefresh() {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        Path disp = AppPaths.resolveResultDispatchTableJsonPath(ui);
        if (dispatchJsonField != null && Files.isRegularFile(disp)) {
            dispatchJsonField.setText(disp.toString());
        }
        reloadMemberCachesAndOperators();
        rebuildPreview();
    }

    /** Mirrors stage-2 artifact refresh: fill {@code member_schedule*.json} sibling path when possible. */
    void tryAutofillMemberJsonFromStage2(String memberSchedulePath) {
        if (memberJsonField == null) {
            return;
        }
        String m = memberSchedulePath != null ? memberSchedulePath.strip() : "";
        if (m.isEmpty()) {
            return;
        }
        Path json = siblingJson(Path.of(m));
        if (json != null && Files.isRegularFile(json)) {
            memberJsonField.setText(json.toString());
            reloadMemberCachesAndOperators();
        }
    }

    private static Path siblingJson(Path workbookPath) {
        Path fn = workbookPath.getFileName();
        if (fn == null) {
            return null;
        }
        String name = fn.toString();
        String stem;
        if (name.endsWith(".xlsx")) {
            stem = name.substring(0, name.length() - 5);
        } else if (name.endsWith(".xlsm")) {
            stem = name.substring(0, name.length() - 5);
        } else if (name.endsWith(".json")) {
            return workbookPath;
        } else {
            return null;
        }
        return workbookPath.resolveSibling(stem + ".json");
    }

    private void reloadMemberCachesAndOperators() {
        String mp = memberJsonField != null ? memberJsonField.getText().strip() : "";
        if (mp.isEmpty()) {
            cachedMemberSheets = Map.of();
            if (operatorCombo != null) {
                operatorCombo.getItems().clear();
            }
            return;
        }
        Path p = Path.of(mp);
        if (!Files.isRegularFile(p)) {
            statusLabel.setText("ファイルなし: " + p);
            return;
        }
        try {
            Map<String, SheetTable> loaded = JsonTableIo.loadSheetsWorkbook(p);
            cachedMemberSheets = loaded;
            applyInferredStartDateIfNeeded();
            List<String> ops = JsonTableIo.memberOperatorNames(cachedMemberSheets);
            if (operatorCombo != null) {
                String prev = operatorCombo.getValue();
                operatorCombo.getItems().setAll(ops);
                if (prev != null && ops.contains(prev)) {
                    operatorCombo.setValue(prev);
                } else if (!ops.isEmpty()) {
                    operatorCombo.setValue(ops.get(0));
                }
            }
            statusLabel.setText(
                    "読み込み: オペレーター " + ops.size() + " 名");
        } catch (Exception ex) {
            statusLabel.setText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
        }
    }

    private SheetTable sampleMemberSheetForDateColumns() {
        if (cachedMemberSheets.isEmpty()) {
            return null;
        }
        String op = operatorCombo != null ? operatorCombo.getValue() : null;
        if (op != null && cachedMemberSheets.containsKey(op)) {
            return cachedMemberSheets.get(op);
        }
        for (SheetTable st : cachedMemberSheets.values()) {
            if (st.columns().contains("時間帯")) {
                return st;
            }
        }
        return cachedMemberSheets.values().iterator().next();
    }

    private void applyInferredStartDateIfNeeded() {
        if (startDateCalendar == null || cachedMemberSheets.isEmpty()) {
            return;
        }
        SheetTable sample = sampleMemberSheetForDateColumns();
        if (sample == null) {
            return;
        }
        LocalDate current =
                startDateCalendar.getSelectedDate() != null
                        ? startDateCalendar.getSelectedDate()
                        : LocalDate.now();
        int days = selectedDayCount();
        if (OperatorCardDocumentBuilder.canResolveDayColumns(sample.columns(), current, days)) {
            return;
        }
        LocalDate inferred =
                OperatorCardDocumentBuilder.inferScheduleStartDate(cachedMemberSheets, current);
        if (!Objects.equals(inferred, current)) {
            startDateCalendar.setSelectedDate(inferred);
        }
    }

    private void ensureMemberCacheLoaded() {
        String mp = memberJsonField != null ? memberJsonField.getText().strip() : "";
        if (cachedMemberSheets.isEmpty() && !mp.isEmpty()) {
            reloadMemberCachesAndOperators();
        }
    }

    private void rebuildPreview() {
        if (previewHost == null) {
            return;
        }
        ensureMemberCacheLoaded();
        applyInferredStartDateIfNeeded();
        try {
            OperatorCardPage page = buildSelectedPage();
            String font = fontCombo != null ? fontCombo.getValue() : "SansSerif";
            Parent root = OperatorCardPreviewFactory.buildRoot(page, font);
            ScrollPane sp = new ScrollPane(root);
            sp.setFitToWidth(true);
            sp.setPannable(true);
            previewHost.getChildren().setAll(sp);
            statusLabel.setText(
                    "プレビュー: " + page.operatorName() + " / " + page.days().size() + " 日分");
        } catch (Exception ex) {
            statusLabel.setText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
            Label err = new Label(String.valueOf(ex.getMessage()));
            previewHost.getChildren().setAll(err);
        }
    }

    private OperatorCardPage buildSelectedPage() throws IOException, OperatorCardBuildException {
        ensureMemberCacheLoaded();
        applyInferredStartDateIfNeeded();
        List<Map<String, String>> dispatchRows = loadDispatchRows();
        LocalDate start =
                startDateCalendar != null ? startDateCalendar.getSelectedDate() : LocalDate.now();
        if (start == null) {
            throw new OperatorCardBuildException("start date is null");
        }
        String op = operatorCombo != null ? operatorCombo.getValue() : null;
        if (op == null || op.isBlank()) {
            throw new OperatorCardBuildException("select operator");
        }
        return OperatorCardDocumentBuilder.buildPage(
                op, cachedMemberSheets, dispatchRows, start, selectedDayCount());
    }

    private List<Map<String, String>> loadDispatchRows() throws IOException {
        String dp = dispatchJsonField != null ? dispatchJsonField.getText().strip() : "";
        if (dp.isEmpty()) {
            throw new IOException("results dispatch JSON path is empty");
        }
        Path p = Path.of(dp);
        if (!Files.isRegularFile(p)) {
            throw new IOException("dispatch file not found: " + p);
        }
        SheetTable t = JsonTableIo.loadFlatTable(p);
        List<Map<String, String>> rows = new ArrayList<>();
        for (Map<String, String> r : t.rows()) {
            rows.add(r);
        }
        return rows;
    }

    private void printCards() {
        if (shell == null) {
            return;
        }
        reloadMemberCachesAndOperators();
        if (cachedMemberSheets.isEmpty()) {
            statusLabel.setText("member_schedule JSON を読み込めません");
            return;
        }
        List<Map<String, String>> dispatchRows;
        try {
            dispatchRows = loadDispatchRows();
        } catch (IOException ex) {
            statusLabel.setText(ex.getMessage());
            return;
        }
        LocalDate start =
                startDateCalendar != null ? startDateCalendar.getSelectedDate() : LocalDate.now();
        if (start == null) {
            statusLabel.setText("開始日を設定してください");
            return;
        }
        List<String> operators = new ArrayList<>();
        boolean all =
                printAllOperatorsCheckBox != null && printAllOperatorsCheckBox.isSelected();
        if (all) {
            operators.addAll(JsonTableIo.memberOperatorNames(cachedMemberSheets));
        } else {
            String op = operatorCombo != null ? operatorCombo.getValue() : null;
            if (op == null || op.isBlank()) {
                statusLabel.setText("オペレーターを選択してください");
                return;
            }
            operators.add(op);
        }
        if (operators.isEmpty()) {
            statusLabel.setText("印刷対象がありません");
            return;
        }

        String font = fontCombo != null ? fontCombo.getValue() : "SansSerif";
        int dayCount = selectedDayCount();

        PrinterJob job = PrinterJob.createPrinterJob();
        if (!job.showPrintDialog(ownerStage)) {
            return;
        }
        Printer printer = job.getPrinter();
        PageLayout layout =
                printer.createPageLayout(
                        Paper.A4, PageOrientation.PORTRAIT, Printer.MarginType.DEFAULT);

        try {
            for (String opName : operators) {
                OperatorCardPage page =
                        OperatorCardDocumentBuilder.buildPage(
                                opName, cachedMemberSheets, dispatchRows, start, dayCount);
                Parent root = OperatorCardPreviewFactory.buildRoot(page, font);
                boolean ok = job.printPage(layout, root);
                if (!ok) {
                    shell.appendLog("[operator-card] printPage returned false for " + opName);
                    break;
                }
            }
        } catch (Exception ex) {
            statusLabel.setText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
            shell.appendLog("[operator-card] " + ex.getMessage());
            return;
        } finally {
            job.endJob();
        }
        statusLabel.setText(
                "印刷完了: " + operators.size() + " 名分");
    }
}
