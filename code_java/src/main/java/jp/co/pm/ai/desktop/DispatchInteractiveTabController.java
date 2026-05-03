package jp.co.pm.ai.desktop;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.Node;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.TabPane;
import javafx.scene.control.TableCell;
import javafx.scene.control.TextInputDialog;
import javafx.scene.control.ToggleButton;
import javafx.scene.control.Tooltip;
import javafx.scene.input.ClipboardContent;
import javafx.scene.input.DragEvent;
import javafx.scene.input.Dragboard;
import javafx.scene.input.MouseEvent;
import javafx.scene.input.TransferMode;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.Window;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetCell;
import org.controlsfx.control.spreadsheet.SpreadsheetCellType;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.dispatch.MachineCalendarBlockIndex;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchDocument;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchJsonIo;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchNormalizer;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchPivot;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchPythonExport;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchSchema;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchTrialPython;
import jp.co.pm.ai.desktop.ui.SpreadsheetTabularSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetThemeBridge;

/**
 * Interactive pivot editor for result dispatch JSON (ControlsFX SpreadsheetView: task-by-day quarters +
 * process+machine-by-day).
 */
public final class DispatchInteractiveTabController {

    private record ReloadBundle(
            ResultDispatchDocument doc,
            MachineCalendarBlockIndex calendar,
            String calendarLoadError,
            String pythonCalendarJsonError,
            String pythonCalendarDiagnosticsJson) {}

    private record CalendarReloadOutcome(
            MachineCalendarBlockIndex calendar,
            String error,
            String pythonCalendarJsonError,
            String pythonCalendarDiagnosticsJson) {}

    private static final String DND_PREFIX = "pm-dispatch-dnd|wide|";
    private static final String DND_V2_MARKER = "v2|";

    private static final int QUARTERS_PER_DAY = 4;

    private static final List<String> WIDE_STATIC_HEADERS =
            List.of(
                    ResultDispatchSchema.COL_DISPATCH_TRIAL_ORDER,
                    ResultDispatchSchema.COL_PROCESS,
                    ResultDispatchSchema.COL_MACHINE,
                    "\u4f9d\u983cNO",
                    "\u63db\u7b97\u6570\u91cf",
                    "\u8a08\u753b\u5408\u8a08");

    @FXML
    private Button loadButton;

    @FXML
    private Button saveButton;

    @FXML
    private Button reloadCalendarButton;

    @FXML
    private Button dispatchTrialButton;

    @FXML
    private Button wideRowUpButton;

    @FXML
    private Button wideRowDownButton;

    @FXML
    private ToggleButton staffCheckToggle;

    @FXML
    private Label statusLabel;

    @FXML
    private ProgressBar reloadProgressBar;

    @FXML
    private Label jsonPathLabel;

    @FXML
    private TabPane innerTabPane;

    @FXML
    private StackPane wideSpreadsheetHost;

    @FXML
    private StackPane byDaySpreadsheetHost;

    private final SpreadsheetView wideSpreadsheet = new SpreadsheetView();
    private final SpreadsheetView byDaySpreadsheet = new SpreadsheetView();

    private MainShellController shell;

    private ResultDispatchDocument doc = ResultDispatchDocument.empty();

    private List<LocalDate> dateAxis = new ArrayList<>();

    private final List<Map<String, String>> wideProfiles = new ArrayList<>();

    /** Parallel to {@link #wideProfiles} rows in the wide grid. */
    private final List<WideRow> wideRowItems = new ArrayList<>();

    private MachineCalendarBlockIndex calendarBlocks = MachineCalendarBlockIndex.empty();

    @FXML
    private void initialize() {
        StackPane.setAlignment(wideSpreadsheet, Pos.CENTER_LEFT);
        wideSpreadsheetHost.getChildren().setAll(wideSpreadsheet);
        VBox.setVgrow(wideSpreadsheetHost, javafx.scene.layout.Priority.ALWAYS);

        StackPane.setAlignment(byDaySpreadsheet, Pos.CENTER_LEFT);
        byDaySpreadsheetHost.getChildren().setAll(byDaySpreadsheet);
        VBox.setVgrow(byDaySpreadsheetHost, javafx.scene.layout.Priority.ALWAYS);

        SpreadsheetThemeBridge.install(wideSpreadsheet);
        SpreadsheetThemeBridge.install(byDaySpreadsheet);

        wideSpreadsheet.getSelectionModel().setSelectionMode(javafx.scene.control.SelectionMode.MULTIPLE);
        byDaySpreadsheet.getSelectionModel().setSelectionMode(javafx.scene.control.SelectionMode.MULTIPLE);

        staffCheckToggle
                .selectedProperty()
                .addListener(
                        (obs, o, n) -> {
                            rebuildGrids();
                        });

        installWideDnDHandlers();
        installByDayDoubleClickHandler();
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        Platform.runLater(this::reloadFromDiskQuiet);
    }

    @FXML
    private void onLoadAction() {
        reloadFromDiskQuiet();
    }

    @FXML
    private void onSaveAction() {
        if (shell == null) {
            return;
        }
        try {
            Path jsonPath = AppPaths.resolveResultDispatchTableJsonPath(shell.snapshotUiEnv());
            ResultDispatchJsonIo.write(jsonPath, doc);
            Path pyExe = resolvePythonExe();
            Path pyDir = AppPaths.resolvePythonScriptDir(shell.snapshotUiEnv());
            String xlsxOut = ResultDispatchPythonExport.exportXlsxNearJson(jsonPath, pyExe, pyDir);
            statusLabel.setText("\u4fdd\u5b58\u3057\u307e\u3057\u305f");
            shell.appendLog("[dispatch-editor] saved json: " + jsonPath);
            if (xlsxOut != null && !xlsxOut.isEmpty()) {
                shell.appendLog("[dispatch-editor] xlsx: " + xlsxOut);
            } else {
                shell.appendLog("[dispatch-editor] xlsx export skipped or failed (Python)");
            }
        } catch (Exception e) {
            statusLabel.setText("\u4fdd\u5b58\u30a8\u30e9\u30fc");
            shell.appendLog("[dispatch-editor] save failed: " + e.getMessage());
        }
    }

    @FXML
    private void onReloadCalendarAction() {
        if (shell == null) {
            return;
        }
        showReloadProgress();
        final MainShellController shellRef = shell;
        Task<CalendarReloadOutcome> task =
                new Task<>() {
                    @Override
                    protected CalendarReloadOutcome call() {
                        try {
                            MachineCalendarBlockIndex.LoadOutcome lo =
                                    loadMachineCalendarFromSharedJson(shellRef);
                            return new CalendarReloadOutcome(
                                    lo.index(),
                                    null,
                                    lo.pythonJsonError(),
                                    lo.pythonDiagnosticsJson());
                        } catch (Exception e) {
                            return new CalendarReloadOutcome(
                                    MachineCalendarBlockIndex.empty(), e.getMessage(), null, null);
                        }
                    }
                };
        task.setOnSucceeded(
                ev -> {
                    CalendarReloadOutcome o = task.getValue();
                    calendarBlocks = o.calendar();
                    if (o.error() != null) {
                        shell.appendLog("[dispatch-editor] calendar load: " + o.error());
                    }
                    if (o.pythonCalendarJsonError() != null) {
                        shell.appendLog(
                                "[dispatch-editor] machine calendar json: "
                                        + o.pythonCalendarJsonError());
                    }
                    if (o.pythonCalendarDiagnosticsJson() != null) {
                        shell.appendLog(
                                "[dispatch-editor] machine calendar diagnostics: "
                                        + o.pythonCalendarDiagnosticsJson());
                    }
                    if (o.error() == null) {
                        shell.appendLog(
                                "[dispatch-editor] machine calendar blocks: "
                                        + (calendarBlocks.isEmpty() ? "none" : "loaded"));
                    }
                    maybeLogMachineCalendarEmptyBlocksHint(
                            o.error(), o.pythonCalendarJsonError(), o.pythonCalendarDiagnosticsJson());
                    rebuildGrids();
                    hideReloadProgress();
                });
        task.setOnFailed(
                ev -> {
                    calendarBlocks = MachineCalendarBlockIndex.empty();
                    Throwable ex = task.getException();
                    shell.appendLog(
                            "[dispatch-editor] calendar load: "
                                    + (ex != null ? ex.getMessage() : ""));
                    rebuildGrids();
                    hideReloadProgress();
                });
        new Thread(task, "dispatch-calendar-reload").start();
    }

    @FXML
    private void onDispatchTrialAction() {
        if (shell == null) {
            return;
        }
        statusLabel.setText("\u914d\u53f0\u8a66\u884c\u4e2d...");
        Path jsonPath = AppPaths.resolveResultDispatchTableJsonPath(shell.snapshotUiEnv());
        Task<String> task =
                new Task<>() {
                    @Override
                    protected String call() throws Exception {
                        Path pyExe = resolvePythonExe();
                        Path pyDir = AppPaths.resolvePythonScriptDir(shell.snapshotUiEnv());
                        return ResultDispatchTrialPython.runTrial(jsonPath, pyExe, pyDir);
                    }
                };
        task.setOnSucceeded(
                e -> {
                    String shortagesPath = task.getValue();
                    statusLabel.setText("\u914d\u53f0\u8a66\u884c\u5b8c\u4e86");
                    shell.appendLog("[dispatch-editor] trial: " + shortagesPath);
                    Alert a = new Alert(Alert.AlertType.INFORMATION);
                    a.initOwner(shell.getPrimaryStage());
                    a.setTitle("\u914d\u53f0\u8a66\u884c");
                    a.setHeaderText(null);
                    a.setContentText(
                            "\u7d50\u679c\u3092\u66f4\u65b0\u3057\u307e\u3057\u305f\u3002\n"
                                    + shortagesPath);
                    a.show();
                    reloadFromDiskQuiet();
                });
        task.setOnFailed(
                e -> {
                    Throwable ex = task.getException();
                    statusLabel.setText("\u914d\u53f0\u8a66\u884c\u30a8\u30e9\u30fc");
                    shell.appendLog(
                            "[dispatch-editor] trial failed: "
                                    + (ex != null ? ex.getMessage() : ""));
                });
        new Thread(task, "dispatch-trial").start();
    }

    @FXML
    private void onWideRowUpAction() {
        int i = selectedWideProfileIndex();
        if (i <= 0) {
            return;
        }
        swapWideProfiles(i - 1, i);
    }

    @FXML
    private void onWideRowDownAction() {
        int i = selectedWideProfileIndex();
        if (i < 0 || i >= wideProfiles.size() - 1) {
            return;
        }
        swapWideProfiles(i, i + 1);
    }

    private void reloadFromDiskQuiet() {
        if (shell == null) {
            return;
        }
        Path p = AppPaths.resolveResultDispatchTableJsonPath(shell.snapshotUiEnv());
        jsonPathLabel.setText(p.toString());
        if (!Files.isRegularFile(p)) {
            statusLabel.setText("\u30d5\u30a1\u30a4\u30eb\u306a\u3057");
            doc = ResultDispatchDocument.empty();
            rebuildGrids();
            return;
        }
        showReloadProgress();
        final MainShellController shellRef = shell;
        final Path jsonPath = p;
        Task<ReloadBundle> task =
                new Task<>() {
                    @Override
                    protected ReloadBundle call() throws Exception {
                        ResultDispatchDocument d = ResultDispatchJsonIo.read(jsonPath);
                        MachineCalendarBlockIndex cal = MachineCalendarBlockIndex.empty();
                        String calErr = null;
                        String pyCalErr = null;
                        String pyDiag = null;
                        try {
                            MachineCalendarBlockIndex.LoadOutcome lo =
                                    loadMachineCalendarFromSharedJson(shellRef);
                            cal = lo.index();
                            pyCalErr = lo.pythonJsonError();
                            pyDiag = lo.pythonDiagnosticsJson();
                        } catch (Exception e) {
                            calErr = e.getMessage();
                        }
                        return new ReloadBundle(d, cal, calErr, pyCalErr, pyDiag);
                    }
                };
        task.setOnSucceeded(
                ev -> {
                    ReloadBundle b = task.getValue();
                    doc = b.doc();
                    calendarBlocks = b.calendar();
                    statusLabel.setText(doc.rows().size() + " \u884c");
                    if (b.calendarLoadError() != null) {
                        shell.appendLog("[dispatch-editor] calendar load: " + b.calendarLoadError());
                    }
                    if (b.pythonCalendarJsonError() != null) {
                        shell.appendLog(
                                "[dispatch-editor] machine calendar json: "
                                        + b.pythonCalendarJsonError());
                    }
                    if (b.pythonCalendarDiagnosticsJson() != null) {
                        shell.appendLog(
                                "[dispatch-editor] machine calendar diagnostics: "
                                        + b.pythonCalendarDiagnosticsJson());
                    }
                    if (b.calendarLoadError() == null) {
                        shell.appendLog(
                                "[dispatch-editor] machine calendar blocks: "
                                        + (calendarBlocks.isEmpty() ? "none" : "loaded"));
                    }
                    maybeLogMachineCalendarEmptyBlocksHint(
                            b.calendarLoadError(),
                            b.pythonCalendarJsonError(),
                            b.pythonCalendarDiagnosticsJson());
                    rebuildGrids();
                    hideReloadProgress();
                });
        task.setOnFailed(
                ev -> {
                    doc = ResultDispatchDocument.empty();
                    calendarBlocks = MachineCalendarBlockIndex.empty();
                    statusLabel.setText("\u8aad\u8fbc\u30a8\u30e9\u30fc");
                    Throwable ex = task.getException();
                    shell.appendLog(
                            "[dispatch-editor] load failed: "
                                    + (ex != null ? ex.getMessage() : ""));
                    rebuildGrids();
                    hideReloadProgress();
                });
        new Thread(task, "dispatch-editor-reload").start();
    }

    private MachineCalendarBlockIndex.LoadOutcome loadMachineCalendarFromSharedJson(MainShellController shellRef)
            throws Exception {
        Path json = AppPaths.resolveMachineCalendarBlocksJsonPath(shellRef.snapshotUiEnv());
        return MachineCalendarBlockIndex.loadOutcomeFromJsonFile(json);
    }

    void refreshCalendarFromSharedJsonFile() {
        if (shell == null) {
            return;
        }
        Runnable r =
                () -> {
                    try {
                        MachineCalendarBlockIndex.LoadOutcome lo =
                                MachineCalendarBlockIndex.loadOutcomeFromJsonFile(
                                        AppPaths.resolveMachineCalendarBlocksJsonPath(shell.snapshotUiEnv()));
                        calendarBlocks = lo.index();
                        rebuildGrids();
                        if ("missing_file".equals(lo.pythonJsonError())) {
                            shell.appendLog(
                                    "[dispatch-editor] machine calendar json missing: "
                                            + AppPaths.resolveMachineCalendarBlocksJsonPath(
                                                    shell.snapshotUiEnv()));
                        }
                    } catch (Exception e) {
                        calendarBlocks = MachineCalendarBlockIndex.empty();
                        rebuildGrids();
                        shell.appendLog(
                                "[dispatch-editor] machine calendar json read: " + e.getMessage());
                    }
                };
        if (Platform.isFxApplicationThread()) {
            r.run();
        } else {
            Platform.runLater(r);
        }
    }

    private void maybeLogMachineCalendarEmptyBlocksHint(
            String javaCalendarLoadError,
            String pythonJsonError,
            String pythonDiagnosticsJson) {
        if (shell == null) {
            return;
        }
        if (javaCalendarLoadError != null) {
            return;
        }
        if ("missing_file".equals(pythonJsonError)) {
            return;
        }
        if (pythonJsonError != null) {
            return;
        }
        if (pythonDiagnosticsJson != null) {
            return;
        }
        if (!calendarBlocks.isEmpty()) {
            return;
        }
        shell.appendLog(
                "[dispatch-editor] machine calendar hint: empty blocks in JSON. Use "
                        + "\u6a5f\u68b0\u30ab\u30ec\u30f3\u30c0\u30fc"
                        + " (JSON) tab "
                        + "\u300c\u30de\u30b9\u30bf\u304b\u3089 JSON \u51fa\u529b\u300d"
                        + " or check master / occupancy.");
    }

    private void showReloadProgress() {
        if (reloadProgressBar != null) {
            reloadProgressBar.setManaged(true);
            reloadProgressBar.setVisible(true);
            reloadProgressBar.setProgress(ProgressBar.INDETERMINATE_PROGRESS);
        }
        setReloadInteractionDisabled(true);
    }

    private void hideReloadProgress() {
        if (reloadProgressBar != null) {
            reloadProgressBar.setProgress(0);
            reloadProgressBar.setVisible(false);
            reloadProgressBar.setManaged(false);
        }
        setReloadInteractionDisabled(false);
    }

    private void setReloadInteractionDisabled(boolean disabled) {
        if (loadButton != null) {
            loadButton.setDisable(disabled);
        }
        if (saveButton != null) {
            saveButton.setDisable(disabled);
        }
        if (reloadCalendarButton != null) {
            reloadCalendarButton.setDisable(disabled);
        }
        if (dispatchTrialButton != null) {
            dispatchTrialButton.setDisable(disabled);
        }
        if (wideRowUpButton != null) {
            wideRowUpButton.setDisable(disabled);
        }
        if (wideRowDownButton != null) {
            wideRowDownButton.setDisable(disabled);
        }
        if (staffCheckToggle != null) {
            staffCheckToggle.setDisable(disabled);
        }
    }

    private Path resolvePythonExeForShell(MainShellController shellRef) {
        if (shellRef == null) {
            return Path.of("python3");
        }
        String py = shellRef.snapshotUiEnv().get(AppPaths.KEY_PM_AI_PYTHON);
        if (py != null && !py.isBlank()) {
            return Path.of(py.trim());
        }
        return Path.of("python3");
    }

    private Path resolvePythonExe() {
        return resolvePythonExeForShell(shell);
    }

    private void rebuildGrids() {
        ensureDateAxis();
        rebuildWide();
        rebuildByDay();
    }

    private void ensureDateAxis() {
        List<LocalDate> distinct = ResultDispatchPivot.distinctDates(doc.rows());
        if (distinct.isEmpty()) {
            dateAxis = new ArrayList<>();
            LocalDate t = LocalDate.now();
            for (int i = 0; i < 14; i++) {
                dateAxis.add(t.plusDays(i));
            }
        } else {
            dateAxis = ResultDispatchPivot.dateRangeInclusive(distinct);
        }
    }

    private void rebuildWide() {
        wideProfiles.clear();
        wideRowItems.clear();
        List<String> cols = doc.columns();
        wideProfiles.addAll(ResultDispatchPivot.distinctTaskProfiles(cols, doc.rows()));
        wideProfiles.sort(
                Comparator.comparing(DispatchInteractiveTabController::parseTrialOrderKey)
                        .thenComparing(p -> ResultDispatchNormalizer.staticGroupKey(cols, p)));
        assignSequentialTrialOrders();

        int staticCols = WIDE_STATIC_HEADERS.size();
        int dayCount = dateAxis.size();
        int slotCols = dayCount * QUARTERS_PER_DAY;
        int totalCols = staticCols + slotCols;
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        int gridRowsTotal = firstData + wideProfiles.size();
        GridBase grid = new GridBase(gridRowsTotal, totalCols);
        grid.getColumnHeaders().clear();
        grid.getColumnHeaders().addAll(buildWideColumnLabels());

        List<ObservableList<SpreadsheetCell>> gridRows = new ArrayList<>(gridRowsTotal);

        ObservableList<SpreadsheetCell> filterRow = FXCollections.observableArrayList();
        for (int c = 0; c < totalCols; c++) {
            SpreadsheetCell cell =
                    SpreadsheetCellType.STRING.createCell(
                            SpreadsheetTabularSupport.SPREADSHEET_FILTER_ROW, c, 1, 1, "");
            cell.setEditable(false);
            filterRow.add(cell);
        }
        gridRows.add(filterRow);

        for (int pr = 0; pr < wideProfiles.size(); pr++) {
            Map<String, String> profile = wideProfiles.get(pr);
            int gridRow = firstData + pr;
            WideRow wr = new WideRow(profile, dateAxis.size());
            for (int j = 0; j < dateAxis.size(); j++) {
                double v =
                        ResultDispatchPivot.sumQuantityForProfileAndDate(
                                cols, doc.rows(), profile, dateAxis.get(j));
                wr.setAmount(j, v);
            }
            wideRowItems.add(wr);

            ObservableList<SpreadsheetCell> line = FXCollections.observableArrayList();
            for (int c = 0; c < staticCols; c++) {
                String title = WIDE_STATIC_HEADERS.get(c);
                String raw = wr.getStatic(title);
                SpreadsheetCell cell =
                        SpreadsheetCellType.STRING.createCell(gridRow, c, 1, 1, raw != null ? raw : "");
                cell.setEditable(c > 0);
                line.add(cell);
            }
            for (int di = 0; di < dayCount; di++) {
                double dayAmt = wr.getAmount(di);
                double q = dayAmt / QUARTERS_PER_DAY;
                String qtxt = dayAmt > 1e-9 ? ResultDispatchNormalizer.formatQty(q) : "";
                for (int qn = 0; qn < QUARTERS_PER_DAY; qn++) {
                    int col = staticCols + di * QUARTERS_PER_DAY + qn;
                    SpreadsheetCell cell =
                            SpreadsheetCellType.STRING.createCell(gridRow, col, 1, 1, qtxt);
                    cell.setEditable(false);
                    applyWideCellStyle(pr, di, cell);
                    line.add(cell);
                }
            }
            gridRows.add(line);
        }
        grid.setRows(gridRows);

        wideSpreadsheet.setGrid(grid);
        wideSpreadsheet.setFilteredRow(SpreadsheetTabularSupport.SPREADSHEET_FILTER_ROW);
        SpreadsheetTabularSupport.applyColumnFiltersWithDialog(wideSpreadsheet);
        SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicy(wideSpreadsheet);
        SpreadsheetTabularSupport.applyFixedLeadingColumnsLater(wideSpreadsheet, WIDE_STATIC_HEADERS.size());
    }

    private List<String> buildWideColumnLabels() {
        List<String> headers = new ArrayList<>(WIDE_STATIC_HEADERS.size() + dateAxis.size() * QUARTERS_PER_DAY);
        headers.addAll(WIDE_STATIC_HEADERS);
        for (LocalDate d : dateAxis) {
            String ds = d.toString();
            headers.add(ds + " Q1");
            headers.add(ds + " Q2");
            headers.add(ds + " Q3");
            headers.add(ds + " Q4");
        }
        return headers;
    }

    private void applyWideCellStyle(int profileRow, int dateIdx, SpreadsheetCell cell) {
        WideRow wr = wideRowItems.get(profileRow);
        double q = wr.getAmount(dateIdx);
        String proc = wr.getStatic(ResultDispatchSchema.COL_PROCESS);
        String mach = wr.getStatic(ResultDispatchSchema.COL_MACHINE);
        LocalDate day = dateAxis.get(dateIdx);
        boolean block = calendarBlocks.isBlockedDay(proc, mach, day);
        if (block) {
            cell.getStyleClass().add("pm-dispatch-blocked-cell");
            cell.setStyle("-fx-background-color: #c8c8c8;");
        } else if (staffCheckToggle.isSelected() && q > 1e-9) {
            cell.setStyle("-fx-background-color: #ffe0e0;");
        } else {
            cell.setStyle("");
        }
    }

    private static double parseTrialOrderKey(Map<String, String> profile) {
        String s = profile.get(ResultDispatchSchema.COL_DISPATCH_TRIAL_ORDER);
        if (s == null || s.isBlank()) {
            return Double.MAX_VALUE;
        }
        try {
            return Double.parseDouble(s.trim().replace(",", ""));
        } catch (NumberFormatException e) {
            return Double.MAX_VALUE;
        }
    }

    /** Ensures each profile row maps to sequential trial order and pushes into doc rows. */
    private void assignSequentialTrialOrders() {
        List<String> cols = doc.columns();
        for (int i = 0; i < wideProfiles.size(); i++) {
            String ord = Integer.toString(i + 1);
            Map<String, String> prof = wideProfiles.get(i);
            prof.put(ResultDispatchSchema.COL_DISPATCH_TRIAL_ORDER, ord);
            for (Map<String, String> row : doc.rows()) {
                if (ResultDispatchPivot.matchesTaskProfile(cols, prof, row)) {
                    row.put(ResultDispatchSchema.COL_DISPATCH_TRIAL_ORDER, ord);
                }
            }
        }
    }

    private void swapWideProfiles(int a, int b) {
        Map<String, String> pa = wideProfiles.get(a);
        Map<String, String> pb = wideProfiles.get(b);
        wideProfiles.set(a, pb);
        wideProfiles.set(b, pa);
        assignSequentialTrialOrders();
        ResultDispatchNormalizer.normalizeInPlace(doc.columns(), doc.rows());
        rebuildGrids();
    }

    private int selectedWideProfileIndex() {
        var cells = wideSpreadsheet.getSelectionModel().getSelectedCells();
        if (cells == null || cells.isEmpty()) {
            return -1;
        }
        int gridRow = cells.getFirst().getRow();
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        int idx = gridRow - firstData;
        if (idx >= 0 && idx < wideProfiles.size()) {
            return idx;
        }
        return -1;
    }

    private void rebuildByDay() {
        List<String> cols = doc.columns();
        List<Map.Entry<String, String>> keys = ResultDispatchPivot.sortedProcessMachineKeys(doc.rows());
        int staticCols = 2;
        int dayCount = dateAxis.size();
        int slotCols = dayCount * QUARTERS_PER_DAY;
        int totalCols = staticCols + slotCols;
        int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
        int gridRowsTotal = firstData + keys.size();
        GridBase grid = new GridBase(gridRowsTotal, totalCols);
        grid.getColumnHeaders().clear();
        grid.getColumnHeaders().addAll(buildByDayColumnLabels());

        List<ObservableList<SpreadsheetCell>> gridRows = new ArrayList<>(gridRowsTotal);

        ObservableList<SpreadsheetCell> filterRow = FXCollections.observableArrayList();
        for (int c = 0; c < totalCols; c++) {
            SpreadsheetCell cell =
                    SpreadsheetCellType.STRING.createCell(
                            SpreadsheetTabularSupport.SPREADSHEET_FILTER_ROW, c, 1, 1, "");
            cell.setEditable(false);
            filterRow.add(cell);
        }
        gridRows.add(filterRow);

        List<ByDayRow> byItems = new ArrayList<>();
        for (Map.Entry<String, String> en : keys) {
            ByDayRow br = new ByDayRow(en.getKey(), en.getValue(), dateAxis.size());
            for (int j = 0; j < dateAxis.size(); j++) {
                double v =
                        ResultDispatchPivot.sumQuantityForProcessMachineDate(
                                doc.rows(), en.getKey(), en.getValue(), dateAxis.get(j));
                br.setAmount(j, v);
            }
            byItems.add(br);
        }

        for (int ir = 0; ir < byItems.size(); ir++) {
            ByDayRow br = byItems.get(ir);
            int gridRow = firstData + ir;
            ObservableList<SpreadsheetCell> line = FXCollections.observableArrayList();
            SpreadsheetCell c0 =
                    SpreadsheetCellType.STRING.createCell(gridRow, 0, 1, 1, br.process());
            c0.setEditable(false);
            line.add(c0);
            SpreadsheetCell c1 =
                    SpreadsheetCellType.STRING.createCell(gridRow, 1, 1, 1, br.machine());
            c1.setEditable(false);
            line.add(c1);
            for (int di = 0; di < dayCount; di++) {
                double dayAmt = br.getAmount(di);
                double q = dayAmt / QUARTERS_PER_DAY;
                String qtxt = dayAmt > 1e-9 ? ResultDispatchNormalizer.formatQty(q) : "";
                for (int qn = 0; qn < QUARTERS_PER_DAY; qn++) {
                    int col = staticCols + di * QUARTERS_PER_DAY + qn;
                    SpreadsheetCell cell =
                            SpreadsheetCellType.STRING.createCell(gridRow, col, 1, 1, qtxt);
                    cell.setEditable(false);
                    applyByDayCellStyle(br, di, cell);
                    line.add(cell);
                }
            }
            gridRows.add(line);
        }
        grid.setRows(gridRows);

        byDaySpreadsheet.setGrid(grid);
        byDaySpreadsheet.setFilteredRow(SpreadsheetTabularSupport.SPREADSHEET_FILTER_ROW);
        SpreadsheetTabularSupport.applyColumnFiltersWithDialog(byDaySpreadsheet);
        SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicy(byDaySpreadsheet);
        SpreadsheetTabularSupport.applyFixedLeadingColumnsLater(byDaySpreadsheet, 2);
    }

    private List<String> buildByDayColumnLabels() {
        List<String> headers = new ArrayList<>(2 + dateAxis.size() * QUARTERS_PER_DAY);
        headers.add(ResultDispatchSchema.COL_PROCESS);
        headers.add(ResultDispatchSchema.COL_MACHINE);
        for (LocalDate d : dateAxis) {
            String ds = d.toString();
            headers.add(ds + " Q1");
            headers.add(ds + " Q2");
            headers.add(ds + " Q3");
            headers.add(ds + " Q4");
        }
        return headers;
    }

    private void applyByDayCellStyle(ByDayRow br, int dateIdx, SpreadsheetCell cell) {
        double q = br.getAmount(dateIdx);
        boolean block = calendarBlocks.isBlockedDay(br.process(), br.machine(), dateAxis.get(dateIdx));
        if (block) {
            cell.setStyle("-fx-background-color: #c8c8c8;");
        } else if (staffCheckToggle.isSelected() && q > 1e-9) {
            cell.setStyle("-fx-background-color: #ffe0e0;");
        } else {
            cell.setStyle("");
        }
    }

    private void installWideDnDHandlers() {
        wideSpreadsheet.addEventFilter(
                MouseEvent.DRAG_DETECTED,
                e -> {
                    TableCell<?, ?> tc = findTableCell(e.getPickResult().getIntersectedNode());
                    if (tc == null || !isUnderSpreadsheet(wideSpreadsheet, tc)) {
                        return;
                    }
                    int col = tc.getTableView().getColumns().indexOf(tc.getTableColumn());
                    int staticCols = WIDE_STATIC_HEADERS.size();
                    if (col < staticCols) {
                        return;
                    }
                    int slot = col - staticCols;
                    int dateIdx = slot / QUARTERS_PER_DAY;
                    int row = tc.getIndex();
                    int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
                    int profIdx = row - firstData;
                    if (profIdx < 0 || profIdx >= wideRowItems.size() || dateIdx < 0 || dateIdx >= dateAxis.size()) {
                        return;
                    }
                    WideRow wr = wideRowItems.get(profIdx);
                    double qty = wr.getAmount(dateIdx);
                    if (qty <= 1e-9) {
                        return;
                    }
                    Dragboard db = wideSpreadsheet.startDragAndDrop(TransferMode.MOVE);
                    ClipboardContent cc = new ClipboardContent();
                    List<String> cols = doc.columns();
                    String gk = ResultDispatchNormalizer.staticGroupKey(cols, wr.profileMap());
                    String b64 =
                            Base64.getUrlEncoder()
                                    .withoutPadding()
                                    .encodeToString(gk.getBytes(StandardCharsets.UTF_8));
                    cc.putString(DND_PREFIX + DND_V2_MARKER + b64 + ":" + dateIdx + ":" + qty);
                    db.setContent(cc);
                    e.consume();
                });

        wideSpreadsheet.addEventFilter(
                DragEvent.DRAG_OVER,
                e -> {
                    TableCell<?, ?> tc = findTableCell(e.getPickResult().getIntersectedNode());
                    if (tc == null || !isUnderSpreadsheet(wideSpreadsheet, tc)) {
                        return;
                    }
                    int col = tc.getTableView().getColumns().indexOf(tc.getTableColumn());
                    int staticCols = WIDE_STATIC_HEADERS.size();
                    if (col < staticCols) {
                        return;
                    }
                    int slot = col - staticCols;
                    int dateIdx = slot / QUARTERS_PER_DAY;
                    int row = tc.getIndex();
                    int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
                    int profIdx = row - firstData;
                    if (profIdx < 0 || profIdx >= wideRowItems.size()) {
                        return;
                    }
                    WideRow tgt = wideRowItems.get(profIdx);
                    boolean blocked =
                            calendarBlocks.isBlockedDay(
                                    tgt.getStatic(ResultDispatchSchema.COL_PROCESS),
                                    tgt.getStatic(ResultDispatchSchema.COL_MACHINE),
                                    dateAxis.get(dateIdx));
                    if (e.getDragboard().hasString()
                            && e.getDragboard().getString().startsWith(DND_PREFIX)
                            && !blocked) {
                        e.acceptTransferModes(TransferMode.MOVE);
                    }
                    e.consume();
                });

        wideSpreadsheet.addEventFilter(
                DragEvent.DRAG_DROPPED,
                e -> {
                    TableCell<?, ?> tc = findTableCell(e.getPickResult().getIntersectedNode());
                    if (tc == null || !isUnderSpreadsheet(wideSpreadsheet, tc)) {
                        return;
                    }
                    int col = tc.getTableView().getColumns().indexOf(tc.getTableColumn());
                    int staticCols = WIDE_STATIC_HEADERS.size();
                    if (col < staticCols) {
                        return;
                    }
                    int slot = col - staticCols;
                    int dateIdx = slot / QUARTERS_PER_DAY;
                    int row = tc.getIndex();
                    int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
                    int profIdx = row - firstData;
                    if (profIdx < 0 || profIdx >= wideRowItems.size()) {
                        return;
                    }
                    WideRow tgt = wideRowItems.get(profIdx);
                    boolean ok = handleWideDrop(e.getDragboard().getString(), tgt, dateIdx);
                    e.setDropCompleted(ok);
                    e.consume();
                });
    }

    private void installByDayDoubleClickHandler() {
        byDaySpreadsheet.addEventFilter(
                MouseEvent.MOUSE_CLICKED,
                e -> {
                    if (e.getClickCount() != 2) {
                        return;
                    }
                    TableCell<?, ?> tc = findTableCell(e.getPickResult().getIntersectedNode());
                    if (tc == null || !isUnderSpreadsheet(byDaySpreadsheet, tc)) {
                        return;
                    }
                    int col = tc.getTableView().getColumns().indexOf(tc.getTableColumn());
                    int staticCols = 2;
                    if (col < staticCols) {
                        return;
                    }
                    int slot = col - staticCols;
                    int dateIdx = slot / QUARTERS_PER_DAY;
                    int row = tc.getIndex();
                    int firstData = SpreadsheetTabularSupport.spreadsheetFirstDataRowIndex();
                    int dataIdx = row - firstData;
                    List<Map.Entry<String, String>> keys = ResultDispatchPivot.sortedProcessMachineKeys(doc.rows());
                    if (dataIdx < 0 || dataIdx >= keys.size()) {
                        return;
                    }
                    Map.Entry<String, String> en = keys.get(dataIdx);
                    double cur =
                            ResultDispatchPivot.sumQuantityForProcessMachineDate(
                                    doc.rows(), en.getKey(), en.getValue(), dateAxis.get(dateIdx));
                    TextInputDialog dialog =
                            new TextInputDialog(ResultDispatchNormalizer.formatQty(cur));
                    dialog.initOwner(shell != null ? shell.getPrimaryStage() : null);
                    dialog.setTitle("\u65e5\u5225\u5408\u8a08");
                    dialog.setHeaderText(
                            en.getKey()
                                    + " / "
                                    + en.getValue()
                                    + " / "
                                    + dateAxis.get(dateIdx));
                    Optional<String> ov = dialog.showAndWait();
                    ov.filter(s -> !s.isBlank())
                            .ifPresent(
                                    s -> {
                                        double newTotal = ResultDispatchNormalizer.parseDouble(s);
                                        ResultDispatchPivot.scaleProcessMachineDateToTotal(
                                                doc.columns(),
                                                doc.rows(),
                                                en.getKey(),
                                                en.getValue(),
                                                dateAxis.get(dateIdx),
                                                newTotal);
                                        ResultDispatchNormalizer.normalizeInPlace(doc.columns(), doc.rows());
                                        rebuildGrids();
                                    });
                });
    }

    private static TableCell<?, ?> findTableCell(Node n) {
        while (n != null) {
            if (n instanceof TableCell<?, ?> tc) {
                return tc;
            }
            n = n.getParent();
        }
        return null;
    }

    private static boolean isUnderSpreadsheet(SpreadsheetView spv, Node node) {
        Node n = node;
        while (n != null) {
            if (n == spv) {
                return true;
            }
            n = n.getParent();
        }
        return false;
    }

    private int wideProfileIndexForRow(WideRow row) {
        List<String> cols = doc.columns();
        String gk = ResultDispatchNormalizer.staticGroupKey(cols, row.profileMap());
        return indexOfProfileGroupKey(gk);
    }

    private int indexOfProfileGroupKey(String groupKey) {
        List<String> cols = doc.columns();
        for (int i = 0; i < wideProfiles.size(); i++) {
            if (ResultDispatchNormalizer.staticGroupKey(cols, wideProfiles.get(i)).equals(groupKey)) {
                return i;
            }
        }
        return -1;
    }

    private boolean handleWideDrop(String payload, WideRow targetRow, int targetDateIdx) {
        if (!payload.startsWith(DND_PREFIX)) {
            return false;
        }
        String rest = payload.substring(DND_PREFIX.length());
        boolean payloadIsV2 = rest.startsWith(DND_V2_MARKER);
        List<String> cols = doc.columns();

        int fromRow;
        int fromDateIdx;
        double max;

        if (payloadIsV2) {
            String body = rest.substring(DND_V2_MARKER.length());
            String[] p = body.split(":", 3);
            if (p.length < 3) {
                return false;
            }
            try {
                String gk =
                        new String(Base64.getUrlDecoder().decode(p[0]), StandardCharsets.UTF_8);
                fromDateIdx = Integer.parseInt(p[1]);
                max = Double.parseDouble(p[2]);
                fromRow = indexOfProfileGroupKey(gk);
            } catch (IllegalArgumentException e) {
                return false;
            }
            if (fromRow < 0) {
                return false;
            }
        } else {
            String[] p = rest.split(":");
            if (p.length < 3) {
                return false;
            }
            try {
                fromRow = Integer.parseInt(p[0]);
                fromDateIdx = Integer.parseInt(p[1]);
                max = Double.parseDouble(p[2]);
            } catch (NumberFormatException e) {
                return false;
            }
        }

        int toIdx = wideProfileIndexForRow(targetRow);

        if (fromRow != toIdx) {
            statusLabel.setText(
                    "\u7e26\u65b9\u5411\u3078\u306e\u79fb\u52d5\u306f\u3067\u304d\u307e\u305b\u3093\uff08\u6a2a\u306e\u307f\uff09");
            return false;
        }
        if (fromRow == toIdx && fromDateIdx == targetDateIdx) {
            return false;
        }
        String proc = targetRow.getStatic(ResultDispatchSchema.COL_PROCESS);
        String mach = targetRow.getStatic(ResultDispatchSchema.COL_MACHINE);
        LocalDate tday = dateAxis.get(targetDateIdx);
        if (calendarBlocks.isBlockedDay(proc, mach, tday)) {
            reportMachineCalendarBlockedMoveRejected(tday, proc, mach);
            return false;
        }
        Optional<Double> moved = pickMoveQuantity(shell != null ? shell.getPrimaryStage() : null, max);
        if (moved.isEmpty()) {
            return false;
        }
        double amt = moved.get();
        if (amt <= 1e-9 || amt > max + 1e-9) {
            return false;
        }
        if (fromRow < 0 || fromRow >= wideProfiles.size() || toIdx < 0 || toIdx >= wideProfiles.size()) {
            return false;
        }
        Map<String, String> fromProfile = wideProfiles.get(fromRow);
        Map<String, String> toProfile = wideProfiles.get(toIdx);
        LocalDate fromDay = dateAxis.get(fromDateIdx);
        LocalDate toDay = dateAxis.get(targetDateIdx);

        double fromSum =
                ResultDispatchPivot.sumQuantityForProfileAndDate(cols, doc.rows(), fromProfile, fromDay);
        double toSum = ResultDispatchPivot.sumQuantityForProfileAndDate(cols, doc.rows(), toProfile, toDay);
        ResultDispatchPivot.upsertAllocation(cols, doc.rows(), fromProfile, fromDay, fromSum - amt);
        ResultDispatchPivot.upsertAllocation(cols, doc.rows(), toProfile, toDay, toSum + amt);
        ResultDispatchNormalizer.normalizeInPlace(cols, doc.rows());
        statusLabel.setText("moved");
        rebuildGrids();
        return true;
    }

    private static Optional<Double> pickMoveQuantity(Window owner, double max) {
        TextInputDialog dialog = new TextInputDialog(ResultDispatchNormalizer.formatQty(max));
        dialog.initOwner(owner);
        dialog.setTitle("Quantity");
        dialog.setHeaderText("Amount to move (max " + ResultDispatchNormalizer.formatQty(max) + ")");
        return dialog.showAndWait()
                .map(ResultDispatchNormalizer::parseDouble)
                .filter(v -> v > 0);
    }

    private void reportMachineCalendarBlockedMoveRejected(
            LocalDate day, String process, String machine) {
        String msg =
                "\u6a5f\u68b0\u30ab\u30ec\u30f3\u30c0\u30fc\u3067\u30d6\u30ed\u30c3\u30af\u306e\u305f\u3081\u79fb\u52d5\u3067\u304d\u307e\u305b\u3093: "
                        + day
                        + " / "
                        + process
                        + " / "
                        + machine;
        statusLabel.setText(msg);
        if (shell != null) {
            shell.appendLog("[dispatch-editor] " + msg);
        }
    }

    /** Mutable wide row (amounts indexed by {@link #dateAxis}). */
    public static final class WideRow {
        private final Map<String, String> staticPart;
        private final double[] amounts;

        WideRow(Map<String, String> staticPart, int nDates) {
            this.staticPart = new LinkedHashMap<>(staticPart);
            this.amounts = new double[nDates];
        }

        String getStatic(String col) {
            return staticPart.getOrDefault(col, "");
        }

        Map<String, String> profileMap() {
            return new LinkedHashMap<>(staticPart);
        }

        double getAmount(int di) {
            return amounts[di];
        }

        void setAmount(int di, double v) {
            amounts[di] = v;
        }
    }

    public record ByDayRow(String process, String machine, double[] amounts) {
        ByDayRow(String process, String machine, int n) {
            this(process, machine, new double[n]);
        }

        double getAmount(int i) {
            return amounts[i];
        }

        void setAmount(int i, double v) {
            amounts[i] = v;
        }
    }
}
