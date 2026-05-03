package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import javafx.application.Platform;
import javafx.beans.property.SimpleStringProperty;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TabPane;
import javafx.scene.control.TextInputDialog;
import javafx.scene.control.ToggleButton;
import javafx.scene.control.Tooltip;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.scene.input.ClipboardContent;
import javafx.scene.input.Dragboard;
import javafx.scene.input.TransferMode;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.dispatch.MachineCalendarBlockIndex;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchDocument;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchJsonIo;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchNormalizer;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchPivot;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchPythonExport;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchSchema;

/**
 * Interactive pivot editor for result dispatch JSON (wide task\u00d7day + aggregate process+machine\u00d7day).
 */
public final class DispatchInteractiveTabController {

    private static final String DND_PREFIX = "pm-dispatch-dnd|wide|";

    private static final List<String> WIDE_STATIC_DISPLAY =
            List.of(
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
    private ToggleButton staffCheckToggle;

    @FXML
    private Label statusLabel;

    @FXML
    private Label jsonPathLabel;

    @FXML
    private TabPane innerTabPane;

    @FXML
    private TableView<WideRow> wideTable;

    @FXML
    private TableView<ByDayRow> byDayTable;

    private MainShellController shell;

    private ResultDispatchDocument doc = ResultDispatchDocument.empty();

    private List<LocalDate> dateAxis = new ArrayList<>();

    private final List<Map<String, String>> wideProfiles = new ArrayList<>();

    private MachineCalendarBlockIndex calendarBlocks = MachineCalendarBlockIndex.empty();

    @FXML
    private void initialize() {
        wideTable.setPlaceholder(new Label("Load 7d50_914d53f0.json (read)"));
        byDayTable.setPlaceholder(new Label("Same"));
        staffCheckToggle
                .selectedProperty()
                .addListener(
                        (obs, o, n) -> {
                            wideTable.refresh();
                            byDayTable.refresh();
                        });
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
        reloadCalendarBlocks();
        rebuildGrids();
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
        try {
            doc = ResultDispatchJsonIo.read(p);
            statusLabel.setText(doc.rows().size() + " \u884c");
            reloadCalendarBlocks();
            rebuildGrids();
        } catch (Exception e) {
            statusLabel.setText("\u8aad\u8fbc\u30a8\u30e9\u30fc");
            shell.appendLog("[dispatch-editor] load failed: " + e.getMessage());
        }
    }

    private void reloadCalendarBlocks() {
        if (shell == null) {
            return;
        }
        try {
            Path master =
                    AppPaths.resolveMasterWorkbookPathResolved(
                            shell.snapshotUiEnv(), shell.effectiveTaskInputWorkbookPathForShell());
            Path py = resolvePythonExe();
            Path dir = AppPaths.resolvePythonScriptDir(shell.snapshotUiEnv());
            calendarBlocks = MachineCalendarBlockIndex.load(master, py, dir);
            shell.appendLog(
                    "[dispatch-editor] machine calendar blocks: "
                            + (calendarBlocks.isEmpty() ? "none" : "loaded"));
        } catch (Exception e) {
            calendarBlocks = MachineCalendarBlockIndex.empty();
            shell.appendLog("[dispatch-editor] calendar load: " + e.getMessage());
        }
    }

    private Path resolvePythonExe() {
        if (shell == null) {
            return Path.of("python3");
        }
        String py = shell.snapshotUiEnv().get(AppPaths.KEY_PM_AI_PYTHON);
        if (py != null && !py.isBlank()) {
            return Path.of(py.trim());
        }
        return Path.of("python3");
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
        wideTable.getColumns().clear();
        wideTable.getItems().clear();
        wideProfiles.clear();
        List<String> cols = doc.columns();
        wideProfiles.addAll(ResultDispatchPivot.distinctTaskProfiles(cols, doc.rows()));

        for (String h : WIDE_STATIC_DISPLAY) {
            TableColumn<WideRow, String> col = new TableColumn<>(h);
            final String hc = h;
            col.setCellValueFactory(
                    cd -> new SimpleStringProperty(cd.getValue() != null ? cd.getValue().getStatic(hc) : ""));
            col.setPrefWidth(96);
            wideTable.getColumns().add(col);
        }

        int di = 0;
        for (LocalDate d : dateAxis) {
            final int dateIdx = di++;
            TableColumn<WideRow, String> dc = new TableColumn<>(d.toString());
            dc.setUserData(dateIdx);
            dc.setPrefWidth(88);
            dc.setCellFactory(tc -> newWideDateCell(dateIdx));
            dc.setCellValueFactory(
                    cd ->
                            new SimpleStringProperty(
                                    cd.getValue() != null
                                            ? ResultDispatchNormalizer.formatQty(cd.getValue().getAmount(dateIdx))
                                            : ""));
            wideTable.getColumns().add(dc);
        }

        List<WideRow> items = new ArrayList<>();
        for (Map<String, String> profile : wideProfiles) {
            WideRow wr = new WideRow(profile, dateAxis.size());
            for (int j = 0; j < dateAxis.size(); j++) {
                double v =
                        ResultDispatchPivot.sumQuantityForProfileAndDate(
                                cols, doc.rows(), profile, dateAxis.get(j));
                wr.setAmount(j, v);
            }
            items.add(wr);
        }
        wideTable.getItems().addAll(items);
    }

    private TableCell<WideRow, String> newWideDateCell(int dateIdx) {
        TableCell<WideRow, String> cell =
                new TableCell<>() {
                    @Override
                    protected void updateItem(String item, boolean empty) {
                        super.updateItem(item, empty);
                        if (empty || getTableRow() == null || getTableRow().getItem() == null) {
                            setText(null);
                            setTooltip(null);
                            setStyle("");
                            return;
                        }
                        WideRow wr = getTableRow().getItem();
                        double q = wr.getAmount(dateIdx);
                        String txt = q > 1e-9 ? ResultDispatchNormalizer.formatQty(q) : "";
                        setText(txt);
                        applyCellStyle(this, wr, dateIdx, q);
                        installStaffTooltip(this, wr, dateIdx);
                    }
                };
        cell.setOnDragDetected(
                ev -> {
                    WideRow wr = cell.getTableRow() != null ? cell.getTableRow().getItem() : null;
                    if (wr == null) {
                        return;
                    }
                    double qty = wr.getAmount(dateIdx);
                    if (qty <= 1e-9) {
                        return;
                    }
                    Dragboard db = cell.startDragAndDrop(TransferMode.MOVE);
                    ClipboardContent cc = new ClipboardContent();
                    int rowIdx = wideTable.getItems().indexOf(wr);
                    cc.putString(DND_PREFIX + rowIdx + ":" + dateIdx + ":" + qty);
                    db.setContent(cc);
                    ev.consume();
                });
        cell.setOnDragOver(
                ev -> {
                    if (ev.getGestureSource() != cell && ev.getDragboard().hasString()) {
                        String s = ev.getDragboard().getString();
                        if (s.startsWith(DND_PREFIX)) {
                            ev.acceptTransferModes(TransferMode.MOVE);
                        }
                    }
                    ev.consume();
                });
        cell.setOnDragDropped(
                ev -> {
                    Dragboard db = ev.getDragboard();
                    boolean ok = false;
                    try {
                        WideRow tgt =
                                cell.getTableRow() != null ? cell.getTableRow().getItem() : null;
                        if (db.hasString() && tgt != null) {
                            ok = handleWideDrop(db.getString(), tgt, dateIdx);
                        }
                    } finally {
                        ev.setDropCompleted(ok);
                    }
                    ev.consume();
                });
        return cell;
    }

    private void applyCellStyle(TableCell<WideRow, String> cell, WideRow wr, int dateIdx, double q) {
        String proc = wr.getStatic(ResultDispatchSchema.COL_PROCESS);
        String mach = wr.getStatic(ResultDispatchSchema.COL_MACHINE);
        LocalDate day = dateAxis.get(dateIdx);
        boolean block = calendarBlocks.isBlockedDay(proc, mach, day);
        if (block) {
            cell.setStyle("-fx-background-color: #c8c8c8;");
        } else if (staffCheckToggle.isSelected() && q > 1e-9) {
            cell.setStyle("-fx-background-color: #ffe0e0;");
        } else {
            cell.setStyle("");
        }
    }

    private void installStaffTooltip(TableCell<WideRow, String> cell, WideRow wr, int dateIdx) {
        if (!staffCheckToggle.isSelected()) {
            cell.setTooltip(null);
            return;
        }
        String proc = wr.getStatic(ResultDispatchSchema.COL_PROCESS);
        String mach = wr.getStatic(ResultDispatchSchema.COL_MACHINE);
        LocalDate day = dateAxis.get(dateIdx);
        String msg =
                "Staffing detail (OP/AS): planned JSON integration.\n"
                        + proc
                        + " / "
                        + mach
                        + " / "
                        + day;
        cell.setTooltip(new Tooltip(msg));
    }

    private boolean handleWideDrop(String payload, WideRow targetRow, int targetDateIdx) {
        if (!payload.startsWith(DND_PREFIX)) {
            return false;
        }
        String rest = payload.substring(DND_PREFIX.length());
        String[] p = rest.split(":");
        if (p.length < 3) {
            return false;
        }
        int fromRow = Integer.parseInt(p[0]);
        int fromDateIdx = Integer.parseInt(p[1]);
        double max = Double.parseDouble(p[2]);
        int toIdx = wideTable.getItems().indexOf(targetRow);
        if (fromRow == toIdx && fromDateIdx == targetDateIdx) {
            return false;
        }
        String proc = targetRow.getStatic(ResultDispatchSchema.COL_PROCESS);
        String mach = targetRow.getStatic(ResultDispatchSchema.COL_MACHINE);
        LocalDate tday = dateAxis.get(targetDateIdx);
        if (calendarBlocks.isBlockedDay(proc, mach, tday)) {
            statusLabel.setText("\u6a5f\u68b0\u30ab\u30ec\u30f3\u30c0\u30fc\u3067\u30d6\u30ed\u30c3\u30af");
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
        List<String> cols = doc.columns();
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

    private void rebuildByDay() {
        byDayTable.getColumns().clear();
        byDayTable.getItems().clear();
        TableColumn<ByDayRow, String> cProc = new TableColumn<>(ResultDispatchSchema.COL_PROCESS);
        cProc.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue().process()));
        cProc.setPrefWidth(120);
        TableColumn<ByDayRow, String> cMach = new TableColumn<>(ResultDispatchSchema.COL_MACHINE);
        cMach.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue().machine()));
        cMach.setPrefWidth(160);
        byDayTable.getColumns().add(cProc);
        byDayTable.getColumns().add(cMach);

        List<String> cols = doc.columns();
        int di = 0;
        for (LocalDate d : dateAxis) {
            final int dateIdx = di++;
            TableColumn<ByDayRow, String> dc = new TableColumn<>(d.toString());
            dc.setPrefWidth(80);
            dc.setCellValueFactory(
                    cd ->
                            new SimpleStringProperty(qtyLabel(cd.getValue().getAmount(dateIdx))));
            dc.setCellFactory(TextFieldTableCell.forTableColumn());
            dc.setOnEditCommit(
                    ev -> {
                        ByDayRow br = ev.getRowValue();
                        double newTotal =
                                ResultDispatchNormalizer.parseDouble(
                                        ev.getNewValue() != null ? ev.getNewValue() : "");
                        ResultDispatchPivot.scaleProcessMachineDateToTotal(
                                cols,
                                doc.rows(),
                                br.process(),
                                br.machine(),
                                dateAxis.get(dateIdx),
                                newTotal);
                        ResultDispatchNormalizer.normalizeInPlace(cols, doc.rows());
                        rebuildGrids();
                    });
            dc.setEditable(true);
            byDayTable.getColumns().add(dc);
        }

        List<Map.Entry<String, String>> keys = ResultDispatchPivot.sortedProcessMachineKeys(doc.rows());
        List<ByDayRow> items = new ArrayList<>();
        for (Map.Entry<String, String> en : keys) {
            ByDayRow br = new ByDayRow(en.getKey(), en.getValue(), dateAxis.size());
            for (int j = 0; j < dateAxis.size(); j++) {
                double v =
                        ResultDispatchPivot.sumQuantityForProcessMachineDate(
                                doc.rows(), en.getKey(), en.getValue(), dateAxis.get(j));
                br.setAmount(j, v);
            }
            items.add(br);
        }
        byDayTable.getItems().addAll(items);
        byDayTable.setEditable(true);
    }

    private static String qtyLabel(double q) {
        return q > 1e-9 ? ResultDispatchNormalizer.formatQty(q) : "";
    }

    /** Mutable wide row for TableView (amounts indexed by {@link #dateAxis}). */
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
