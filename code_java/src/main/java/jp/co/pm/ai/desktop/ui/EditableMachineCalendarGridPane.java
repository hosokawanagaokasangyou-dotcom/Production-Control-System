package jp.co.pm.ai.desktop.ui;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Deque;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.Consumer;

import com.fasterxml.jackson.databind.JsonNode;

import javafx.application.Platform;
import javafx.geometry.HPos;
import javafx.geometry.Pos;
import javafx.geometry.VPos;
import javafx.scene.Node;
import javafx.scene.control.Button;
import javafx.scene.control.ContextMenu;
import javafx.scene.control.Label;
import javafx.scene.control.MenuItem;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.SeparatorMenuItem;
import javafx.scene.control.Tooltip;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.ColumnConstraints;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.scene.shape.Polygon;
import javafx.stage.Stage;
import javafx.stage.Window;

/** 1日分の機械カレンダー（時刻行×設備列）。 */
public final class EditableMachineCalendarGridPane extends VBox {

    private static final DateTimeFormatter SLOT_LABEL =
            DateTimeFormatter.ofPattern("HH:mm");
    private static final int MAX_UNDO_DEPTH = 24;
    private static final double DRAG_THRESHOLD_PX = 2.0;

    private final GridPane grid = new GridPane();
    private final ScrollPane scroll = new ScrollPane(grid);
    private final StackPane scrollHost = new StackPane();
    private final AttendanceGridLoadingOverlay loadingOverlay =
            new AttendanceGridLoadingOverlay("pm-machine-calendar-grid-loading-overlay");
    private final Map<String, Button> cellUiMap = new HashMap<>();
    private final Map<String, CellUi> cellUiByKey = new HashMap<>();

    private LocalDate day;
    private List<ColumnDef> columns = List.of();
    private List<RowDef> rows = new ArrayList<>();
    private Map<String, Map<String, String>> cells = new HashMap<>();
    private Map<String, Map<String, String>> comments = new HashMap<>();
    private final Set<String> commentEditedKeys = new HashSet<>();
    private Window commentDialogOwner;
    private int cellSizePx = AttendanceGridCellSizing.DEFAULT_CELL_PX;
    private int columnWidthPx = AttendanceGridCellSizing.DEFAULT_MACHINE_CALENDAR_COLUMN_PX;
    private int columnGapPx = AttendanceGridCellSizing.DEFAULT_MACHINE_CALENDAR_COLUMN_GAP_PX;
    /** ビューポートに合わせた実レイアウト用セル寸法（{@link #cellSizePx} を上限とする）。 */
    private int layoutCellSizePx = AttendanceGridCellSizing.DEFAULT_CELL_PX;
    private boolean viewportFitScheduled = false;
    private MachineCalendarCellValues.OccupancyMode paintMode =
            MachineCalendarCellValues.OccupancyMode.OCCUPIED;
    private final Deque<GridSnapshot> undoStack = new ArrayDeque<>();
    private boolean dragGestureActive = false;
    private boolean dragMoved = false;
    private final Set<CellCoord> dragVisited = new HashSet<>();
    private double dragAnchorX;
    private double dragAnchorY;
    private Consumer<Boolean> dirtyListener;
    private Consumer<Boolean> undoStateListener;
    private Map<String, Object> savedBaseline = Map.of();
    private boolean gridLoading = false;

    public record ColumnDef(String equipmentKey, String process, String machine) {}

    public record RowDef(String slotIso) {}

    public record CellCoord(String slotIso, String equipmentKey) {}

    private record CellUi(Button button, Polygon commentMark) {}

    private record GridSnapshot(
            Map<String, Map<String, String>> cells,
            Map<String, Map<String, String>> comments,
            Set<String> commentEditedKeys) {}

    public EditableMachineCalendarGridPane() {
        getStyleClass().add("pm-machine-calendar-grid");
        setSpacing(6);

        HBox legendChips =
                new HBox(
                        6,
                        legendChip("· 稼働可", "pm-machine-calendar-legend-available"),
                        legendChip("* 非稼働", "pm-machine-calendar-legend-occupied"));
        legendChips.getStyleClass().add("pm-attendance-legend-chips");
        Label legend =
                new Label(
                        "クリック: 反転　｜ドラッグ: 塗り（ツールバーのモード）　｜見出しクリック: 列/行一括　｜"
                                + "右クリック: コメント（▲=あり）　｜保存で JSON 正本へ反映");
        legend.getStyleClass().add("pm-machine-calendar-grid-legend");
        legend.setWrapText(true);

        scroll.setFitToHeight(false);
        scroll.setMinHeight(120);
        grid.setHgap(columnGapPx);
        grid.setVgap(0);
        scrollHost.getChildren().addAll(scroll, loadingOverlay);
        scrollHost.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
        scrollHost.setMouseTransparent(true);
        VBox.setVgrow(scrollHost, Priority.ALWAYS);
        setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
        scroll.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
        scroll.viewportBoundsProperty().addListener((obs, o, n) -> scheduleViewportFit());
        heightProperty().addListener((obs, o, n) -> scheduleViewportFit());

        installGridGestureFilters();

        getChildren().addAll(legendChips, legend, scrollHost);
    }

    public void setGridLoading(boolean loading) {
        setGridLoading(loading, null);
    }

    public void setGridLoading(boolean loading, String message) {
        gridLoading = loading;
        if (loading) {
            loadingOverlay.setLoading(true, message);
            scroll.setDisable(true);
            toggleStyleClass(this, "pm-machine-calendar-grid-loading", true);
            toggleStyleClass(scrollHost, "pm-machine-calendar-grid-loading", true);
            scrollHost.setMouseTransparent(false);
        } else {
            loadingOverlay.setLoading(false);
            scroll.setDisable(false);
            toggleStyleClass(this, "pm-machine-calendar-grid-loading", false);
            toggleStyleClass(scrollHost, "pm-machine-calendar-grid-loading", false);
            scrollHost.setMouseTransparent(true);
        }
    }

    public void setGridLoadingMessage(String message) {
        if (!gridLoading) {
            return;
        }
        loadingOverlay.setMessage(message);
    }

    public boolean isGridLoading() {
        return gridLoading;
    }

    private static void toggleStyleClass(javafx.scene.Node node, String styleClass, boolean add) {
        if (add) {
            if (!node.getStyleClass().contains(styleClass)) {
                node.getStyleClass().add(styleClass);
            }
        } else {
            node.getStyleClass().remove(styleClass);
        }
    }

    public void setCommentDialogOwner(Window owner) {
        this.commentDialogOwner = owner;
    }

    public void setDirtyListener(Consumer<Boolean> listener) {
        this.dirtyListener = listener;
    }

    public void setUndoStateListener(Consumer<Boolean> listener) {
        this.undoStateListener = listener;
    }

    public void setPaintMode(MachineCalendarCellValues.OccupancyMode mode) {
        if (mode != null) {
            paintMode = mode;
        }
    }

    public MachineCalendarCellValues.OccupancyMode paintMode() {
        return paintMode;
    }

    public boolean canUndo() {
        return !undoStack.isEmpty();
    }

    public void undo() {
        if (undoStack.isEmpty()) {
            return;
        }
        GridSnapshot snap = undoStack.pop();
        cells = deepCopyCells(snap.cells());
        comments = deepCopyCells(snap.comments());
        commentEditedKeys.clear();
        commentEditedKeys.addAll(snap.commentEditedKeys());
        refreshAllCellUi();
        notifyDirty();
        notifyUndoState();
    }

    public void fillAllOccupied() {
        pushUndoSnapshot();
        for (RowDef row : rows) {
            for (ColumnDef col : columns) {
                writeCellValue(new CellCoord(row.slotIso(), col.equipmentKey()),
                        MachineCalendarCellValues.OccupancyMode.OCCUPIED);
            }
        }
        refreshAllCellUi();
        notifyDirty();
    }

    public void clearAll() {
        pushUndoSnapshot();
        for (RowDef row : rows) {
            for (ColumnDef col : columns) {
                writeCellValue(new CellCoord(row.slotIso(), col.equipmentKey()),
                        MachineCalendarCellValues.OccupancyMode.AVAILABLE);
            }
        }
        refreshAllCellUi();
        notifyDirty();
    }

    public void invertAll() {
        pushUndoSnapshot();
        for (RowDef row : rows) {
            for (ColumnDef col : columns) {
                CellCoord coord = new CellCoord(row.slotIso(), col.equipmentKey());
                String next = MachineCalendarCellValues.toggle(readCellValue(coord));
                writeCellValue(
                        coord,
                        MachineCalendarCellValues.isOccupied(next)
                                ? MachineCalendarCellValues.OccupancyMode.OCCUPIED
                                : MachineCalendarCellValues.OccupancyMode.AVAILABLE);
            }
        }
        refreshAllCellUi();
        notifyDirty();
    }

    public void applyPaintModeToColumn(int columnIndex) {
        if (columnIndex < 0 || columnIndex >= columns.size()) {
            return;
        }
        pushUndoSnapshot();
        ColumnDef col = columns.get(columnIndex);
        for (RowDef row : rows) {
            writeCellValue(new CellCoord(row.slotIso(), col.equipmentKey()), paintMode);
        }
        refreshAllCellUi();
        notifyDirty();
    }

    public void applyPaintModeToRow(int rowIndex) {
        if (rowIndex < 0 || rowIndex >= rows.size()) {
            return;
        }
        pushUndoSnapshot();
        String slot = rows.get(rowIndex).slotIso();
        for (ColumnDef col : columns) {
            writeCellValue(new CellCoord(slot, col.equipmentKey()), paintMode);
        }
        refreshAllCellUi();
        notifyDirty();
    }

    public void setCellSizePx(int px) {
        int clamped = AttendanceGridCellSizing.clamp(px);
        if (cellSizePx == clamped) {
            return;
        }
        cellSizePx = clamped;
        rebuildGrid();
    }

    public void setColumnWidthPx(int px) {
        int clamped = AttendanceGridCellSizing.clampMachineCalendarColumnWidth(px);
        if (columnWidthPx == clamped) {
            return;
        }
        columnWidthPx = clamped;
        rebuildGrid();
    }

    public void setColumnGapPx(int px) {
        int clamped = AttendanceGridCellSizing.clampMachineCalendarColumnGap(px);
        if (columnGapPx == clamped) {
            return;
        }
        columnGapPx = clamped;
        grid.setHgap(clamped);
    }

    public boolean hasUnsavedEdits() {
        return !exportPatchJson().equals(savedBaseline);
    }

    public void captureSavedBaseline() {
        savedBaseline = deepCopyPatch(exportPatchJson());
        undoStack.clear();
        commentEditedKeys.clear();
        notifyDirty();
        notifyUndoState();
    }

    public void clearUnsavedEditFlags() {
        savedBaseline = deepCopyPatch(exportPatchJson());
        undoStack.clear();
        commentEditedKeys.clear();
        notifyDirty();
        notifyUndoState();
    }

    public void loadFromDayGridJson(JsonNode node) {
        try {
            day = LocalDate.parse(node.path("date").asText(LocalDate.now().toString()));
        } catch (Exception e) {
            day = LocalDate.now();
        }
        columns = new ArrayList<>();
        node.path("columns").forEach(
                c ->
                        columns.add(
                                new ColumnDef(
                                        c.path("equipment_key").asText(""),
                                        c.path("process").asText(""),
                                        c.path("machine").asText(""))));
        rows = new ArrayList<>();
        cells.clear();
        comments.clear();
        commentEditedKeys.clear();
        node.path("rows").forEach(
                r -> {
                    String slot = r.path("slot").asText("");
                    if (slot.isBlank()) {
                        return;
                    }
                    rows.add(new RowDef(slot));
                    Map<String, String> row = new HashMap<>();
                    JsonNode cellsNode = r.path("cells");
                    if (cellsNode.isObject()) {
                        cellsNode
                                .fields()
                                .forEachRemaining(
                                        e ->
                                                row.put(
                                                        e.getKey(),
                                                        e.getValue().asText("")));
                    }
                    cells.put(slot, row);
                    Map<String, String> rowComments = new HashMap<>();
                    JsonNode commentsNode = r.path("comments");
                    if (commentsNode.isObject()) {
                        commentsNode
                                .fields()
                                .forEachRemaining(
                                        e ->
                                                rowComments.put(
                                                        e.getKey(),
                                                        e.getValue().asText("")));
                    }
                    if (!rowComments.isEmpty()) {
                        comments.put(slot, rowComments);
                    }
                });
        rebuildGrid();
        captureSavedBaseline();
    }

    public Map<String, Object> exportPatchJson() {
        Map<String, Object> patch = new HashMap<>();
        patch.put("date", day != null ? day.toString() : "");
        List<Map<String, Object>> outRows = new ArrayList<>();
        for (RowDef row : rows) {
            Map<String, Object> rowMap = new HashMap<>();
            rowMap.put("slot", row.slotIso());
            Map<String, String> rowCells = cells.getOrDefault(row.slotIso(), Map.of());
            Map<String, String> rowComments = comments.getOrDefault(row.slotIso(), Map.of());
            Map<String, String> outCells = new HashMap<>();
            Map<String, String> outComments = new HashMap<>();
            for (ColumnDef col : columns) {
                String ek = col.equipmentKey();
                String v = rowCells.get(ek);
                if (v != null && !v.isBlank()) {
                    outCells.put(ek, v);
                }
                String c = rowComments.get(ek);
                String key = cellKey(new CellCoord(row.slotIso(), ek));
                if (c != null && !c.isBlank()) {
                    outComments.put(ek, c);
                } else if (commentEditedKeys.contains(key)) {
                    outComments.put(ek, "");
                }
            }
            rowMap.put("cells", outCells);
            if (!outComments.isEmpty()) {
                rowMap.put("comments", outComments);
            }
            outRows.add(rowMap);
        }
        patch.put("rows", outRows);
        return patch;
    }

    private void installGridGestureFilters() {
        grid.addEventFilter(MouseEvent.MOUSE_DRAGGED, this::onGridMouseDragged);
        grid.addEventFilter(MouseEvent.MOUSE_RELEASED, e -> endDragGesture());
        scroll.addEventFilter(MouseEvent.MOUSE_RELEASED, e -> endDragGesture());
    }

    private void onGridMouseDragged(MouseEvent e) {
        if (!dragGestureActive || !e.isPrimaryButtonDown()) {
            return;
        }
        if (!dragMoved) {
            double dx = e.getSceneX() - dragAnchorX;
            double dy = e.getSceneY() - dragAnchorY;
            if (Math.hypot(dx, dy) >= DRAG_THRESHOLD_PX) {
                dragMoved = true;
            }
        }
        if (!dragMoved) {
            return;
        }
        Node pick = e.getPickResult().getIntersectedNode();
        CellCoord coord = coordFromNode(pick);
        if (coord != null) {
            applyPaintToCell(coord);
        }
    }

    private void beginDragGesture(MouseEvent e) {
        dragGestureActive = true;
        dragMoved = false;
        dragVisited.clear();
        dragAnchorX = e.getSceneX();
        dragAnchorY = e.getSceneY();
        pushUndoSnapshot();
    }

    private void endDragGesture() {
        if (!dragGestureActive) {
            return;
        }
        dragGestureActive = false;
        if (dragMoved) {
            notifyDirty();
        }
        dragVisited.clear();
    }

    private void onCellClicked(CellCoord coord) {
        if (dragMoved) {
            return;
        }
        toggleCell(coord);
        updateCellUi(coord);
        notifyDirty();
    }

    private void toggleCell(CellCoord coord) {
        String next = MachineCalendarCellValues.toggle(readCellValue(coord));
        writeCellValue(
                coord,
                MachineCalendarCellValues.isOccupied(next)
                        ? MachineCalendarCellValues.OccupancyMode.OCCUPIED
                        : MachineCalendarCellValues.OccupancyMode.AVAILABLE);
    }

    private void applyPaintToCell(CellCoord coord) {
        if (!dragVisited.add(coord)) {
            return;
        }
        writeCellValue(coord, paintMode);
        updateCellUi(coord);
    }

    private String readCellValue(CellCoord coord) {
        return cells.getOrDefault(coord.slotIso(), Map.of())
                .getOrDefault(coord.equipmentKey(), "");
    }

    private void writeCellValue(CellCoord coord, MachineCalendarCellValues.OccupancyMode mode) {
        Map<String, String> rowCells = cells.computeIfAbsent(coord.slotIso(), k -> new HashMap<>());
        if (mode == MachineCalendarCellValues.OccupancyMode.OCCUPIED) {
            rowCells.put(coord.equipmentKey(), mode.storedValue());
        } else {
            rowCells.remove(coord.equipmentKey());
        }
    }

    private void pushUndoSnapshot() {
        undoStack.push(
                new GridSnapshot(
                        deepCopyCells(cells),
                        deepCopyCells(comments),
                        new HashSet<>(commentEditedKeys)));
        while (undoStack.size() > MAX_UNDO_DEPTH) {
            undoStack.removeLast();
        }
        notifyUndoState();
    }

    private void notifyUndoState() {
        if (undoStateListener != null) {
            undoStateListener.accept(canUndo());
        }
    }

    private void scheduleViewportFit() {
        if (viewportFitScheduled || rows.isEmpty() || dragGestureActive) {
            return;
        }
        viewportFitScheduled = true;
        Platform.runLater(
                () -> {
                    viewportFitScheduled = false;
                    int next = resolveLayoutCellSizePx();
                    if (next != layoutCellSizePx) {
                        layoutCellSizePx = next;
                        rebuildGrid();
                    }
                });
    }

    private int resolveLayoutCellSizePx() {
        if (rows.isEmpty()) {
            return cellSizePx;
        }
        double viewportH = scroll.getViewportBounds().getHeight();
        if (viewportH <= 0 || !scroll.isVisible()) {
            return cellSizePx;
        }
        int fitted = computeFittedCellPxForViewport(viewportH, rows.size());
        int candidate = Math.min(cellSizePx, fitted);
        if (allRowsFitInViewport(viewportH, rows.size(), candidate)) {
            return candidate;
        }
        return cellSizePx;
    }

    private static boolean allRowsFitInViewport(double viewportHeight, int rowCount, int cellPx) {
        return gridHeightForRows(rowCount, cellPx) <= viewportHeight;
    }

    private static double gridHeightForRows(int rowCount, int cellPx) {
        int rowH = AttendanceGridCellSizing.memberCellHeight(cellPx);
        double vgap = 0.0;
        double extraPadding = 4.0;
        double gaps = vgap * (rowCount + 1);
        return rowH + rowCount * rowH + gaps + extraPadding;
    }

    private static int computeFittedCellPxForViewport(double viewportHeight, int rowCount) {
        for (int px = AttendanceGridCellSizing.MAX_PX;
                px >= AttendanceGridCellSizing.MIN_PX;
                px--) {
            if (gridHeightForRows(rowCount, px) <= viewportHeight) {
                return AttendanceGridCellSizing.clamp(px);
            }
        }
        return AttendanceGridCellSizing.MIN_PX;
    }

    private void rebuildGrid() {
        layoutCellSizePx = resolveLayoutCellSizePx();
        int px = layoutCellSizePx;
        grid.setHgap(columnGapPx);
        grid.getChildren().clear();
        grid.getColumnConstraints().clear();
        cellUiMap.clear();
        cellUiByKey.clear();
        if (columns.isEmpty() || rows.isEmpty()) {
            return;
        }
        ColumnConstraints timeCol = new ColumnConstraints();
        int timeW = AttendanceGridCellSizing.machineCalendarTimeColumnWidth(px);
        timeCol.setMinWidth(timeW);
        timeCol.setPrefWidth(timeW);
        timeCol.setMaxWidth(timeW);
        grid.getColumnConstraints().add(timeCol);
        int dataColW = AttendanceGridCellSizing.machineCalendarDataColumnWidth(columnWidthPx);
        for (ColumnDef col : columns) {
            ColumnConstraints cc = new ColumnConstraints();
            cc.setMinWidth(dataColW);
            cc.setPrefWidth(dataColW);
            cc.setMaxWidth(dataColW);
            grid.getColumnConstraints().add(cc);
        }
        Label timeHeader = new Label("時刻");
        timeHeader.getStyleClass().add("pm-machine-calendar-grid-header");
        AttendanceGridCellSizing.applyHeaderLabel(timeHeader, px);
        grid.add(timeHeader, 0, 0);
        List<MachineCalendarColumnHeaderFormat.Display> headerDisplays =
                MachineCalendarColumnHeaderFormat.formatAll(columns);
        for (int c = 0; c < columns.size(); c++) {
            MachineCalendarColumnHeaderFormat.Display display = headerDisplays.get(c);
            Label h = new Label(display.text());
            h.getStyleClass().add("pm-machine-calendar-grid-header");
            h.getStyleClass().add("pm-machine-calendar-grid-header-interactive");
            AttendanceGridCellSizing.applyHeaderLabel(h, px);
            h.setWrapText(true);
            h.setMaxWidth(dataColW);
            Tooltip.install(
                    h,
                    new Tooltip(
                            display.tooltip()
                                    + "\nクリック: ツールバーのモードでこの機械の全時間帯を一括"));
            GridPane.setHalignment(h, HPos.CENTER);
            GridPane.setValignment(h, VPos.CENTER);
            int columnIndex = c;
            h.setOnMouseClicked(e -> applyPaintModeToColumn(columnIndex));
            grid.add(h, c + 1, 0);
        }
        for (int r = 0; r < rows.size(); r++) {
            RowDef row = rows.get(r);
            Label timeLabel = new Label(formatSlotLabel(row.slotIso()));
            timeLabel.getStyleClass().add("pm-machine-calendar-grid-time");
            timeLabel.getStyleClass().add("pm-machine-calendar-grid-header-interactive");
            AttendanceGridCellSizing.applyMachineCalendarTimeLabel(timeLabel, px);
            GridPane.setHalignment(timeLabel, HPos.CENTER);
            Tooltip.install(
                    timeLabel,
                    new Tooltip("クリック: ツールバーのモードでこの時間帯の全設備を一括"));
            int rowIndex = r;
            timeLabel.setOnMouseClicked(e -> applyPaintModeToRow(rowIndex));
            grid.add(timeLabel, 0, r + 1);
            for (int c = 0; c < columns.size(); c++) {
                ColumnDef col = columns.get(c);
                CellCoord coord = new CellCoord(row.slotIso(), col.equipmentKey());
                StackPane cellWrap = createCellWrap(coord, px, dataColW);
                grid.add(cellWrap, c + 1, r + 1);
            }
        }
    }

    private StackPane createCellWrap(CellCoord coord, int rowCellPx, int dataColW) {
        Button cell = createCellButton(coord, rowCellPx, dataColW);
        Polygon commentMark = buildCommentMark();
        commentMark.setVisible(hasComment(readComment(coord)));
        StackPane wrap = new StackPane(cell, commentMark);
        wrap.setMaxWidth(Double.MAX_VALUE);
        wrap.setMaxHeight(Double.MAX_VALUE);
        CellUi ui = new CellUi(cell, commentMark);
        cellUiByKey.put(cellKey(coord), ui);
        return wrap;
    }

    private Button createCellButton(CellCoord coord, int rowCellPx, int dataColW) {
        String val = readCellValue(coord);
        Button cell = new Button(MachineCalendarCellValues.shortLabel(val));
        cell.getStyleClass().add("pm-machine-calendar-cell");
        cell.setUserData(coord);
        AttendanceGridCellSizing.applyMachineCalendarDataCell(cell, dataColW, rowCellPx);
        GridPane.setFillWidth(cell, true);
        GridPane.setFillHeight(cell, true);
        applyOccupiedStyle(cell, val);
        cell.setOnMousePressed(
                e -> {
                    if (e.isPrimaryButtonDown()) {
                        beginDragGesture(e);
                    }
                });
        cell.setOnMouseEntered(
                e -> {
                    if (dragGestureActive && dragMoved && e.isPrimaryButtonDown()) {
                        applyPaintToCell(coord);
                    }
                });
        cell.setOnMouseClicked(e -> onCellClicked(coord));
        cell.setOnContextMenuRequested(
                e -> {
                    if (dragMoved) {
                        return;
                    }
                    e.consume();
                    showCellContextMenu(cell, coord, e.getScreenX(), e.getScreenY());
                });
        cellUiMap.put(cellKey(coord), cell);
        applyCellTooltip(cell, coord);
        return cell;
    }

    private void updateCellUi(CellCoord coord) {
        CellUi ui = cellUiByKey.get(cellKey(coord));
        Button cell = ui != null ? ui.button() : cellUiMap.get(cellKey(coord));
        if (cell == null) {
            return;
        }
        String val = readCellValue(coord);
        cell.setText(MachineCalendarCellValues.shortLabel(val));
        applyOccupiedStyle(cell, val);
        if (ui != null) {
            ui.commentMark().setVisible(hasComment(readComment(coord)));
        }
        applyCellTooltip(cell, coord);
    }

    private void refreshAllCellUi() {
        for (CellUi ui : cellUiByKey.values()) {
            CellCoord coord = (CellCoord) ui.button().getUserData();
            if (coord != null) {
                updateCellUi(coord);
            }
        }
    }

    private String readComment(CellCoord coord) {
        return comments.getOrDefault(coord.slotIso(), Map.of())
                .getOrDefault(coord.equipmentKey(), "");
    }

    private void writeComment(CellCoord coord, String comment) {
        String norm = comment != null ? comment.strip() : "";
        Map<String, String> row = comments.computeIfAbsent(coord.slotIso(), k -> new HashMap<>());
        if (norm.isEmpty()) {
            row.remove(coord.equipmentKey());
            if (row.isEmpty()) {
                comments.remove(coord.slotIso());
            }
        } else {
            row.put(coord.equipmentKey(), norm);
        }
    }

    private static boolean hasComment(String comment) {
        return comment != null && !comment.isBlank();
    }

    private static Polygon buildCommentMark() {
        Polygon mark = new Polygon(0, 0, 8, 0, 8, 8);
        mark.getStyleClass().add("pm-machine-calendar-cell-comment-mark");
        mark.setMouseTransparent(true);
        mark.setTranslateX(6);
        mark.setTranslateY(-6);
        return mark;
    }

    private void showCellContextMenu(
            Button anchor, CellCoord coord, double screenX, double screenY) {
        boolean commentPresent = hasComment(readComment(coord));
        ContextMenu menu = new ContextMenu();
        MenuItem edit = new MenuItem("コメントを入力…");
        edit.setOnAction(ev -> openCommentDialog(coord));
        MenuItem delete = new MenuItem("コメントを削除");
        delete.setDisable(!commentPresent);
        delete.setOnAction(ev -> applyComment(coord, ""));
        menu.getItems().addAll(edit, delete);
        menu.setOnHidden(ev -> releaseCellFocus(anchor));
        menu.show(anchor, screenX, screenY);
    }

    private void releaseCellFocus(Button cell) {
        if (cell != null && cell.isFocused()) {
            scroll.requestFocus();
        }
    }

    private void openCommentDialog(CellCoord coord) {
        String initial = readComment(coord);
        ColumnDef col =
                columns.stream()
                        .filter(c -> c.equipmentKey().equals(coord.equipmentKey()))
                        .findFirst()
                        .orElse(new ColumnDef(coord.equipmentKey(), "", ""));
        String titleCtx = col.process() + "+" + col.machine();
        if (titleCtx.isBlank() || "+".equals(titleCtx.strip())) {
            titleCtx = coord.equipmentKey();
        }
        Stage owner =
                commentDialogOwner instanceof Stage s
                        ? s
                        : commentDialogOwner != null
                                && commentDialogOwner.getScene() != null
                                && commentDialogOwner.getScene().getWindow() instanceof Stage s2
                                ? s2
                                : null;
        MemberAttendanceCellCommentDialog.show(
                owner,
                titleCtx,
                formatSlotLabel(coord.slotIso()),
                initial,
                text -> applyComment(coord, text));
    }

    private void applyComment(CellCoord coord, String comment) {
        String norm = comment != null ? comment.strip() : "";
        if (norm.equals(readComment(coord).strip())) {
            CellUi ui = cellUiByKey.get(cellKey(coord));
            if (ui != null) {
                releaseCellFocus(ui.button());
            }
            return;
        }
        pushUndoSnapshot();
        commentEditedKeys.add(cellKey(coord));
        writeComment(coord, norm);
        updateCellUi(coord);
        notifyDirty();
        CellUi ui = cellUiByKey.get(cellKey(coord));
        if (ui != null) {
            releaseCellFocus(ui.button());
        }
    }

    private void applyCellTooltip(Button cell, CellCoord coord) {
        String val = readCellValue(coord);
        String cmt = readComment(coord);
        if (!hasComment(cmt)) {
            if (MachineCalendarCellValues.isOccupied(val)) {
                cell.setTooltip(new Tooltip("非稼働"));
            } else {
                cell.setTooltip(new Tooltip("稼働可"));
            }
            return;
        }
        StringBuilder tip = new StringBuilder();
        if (MachineCalendarCellValues.isOccupied(val)) {
            tip.append("非稼働");
        } else {
            tip.append("稼働可");
        }
        tip.append("\nコメント: ").append(cmt.trim());
        cell.setTooltip(new Tooltip(tip.toString()));
    }

    private static void applyOccupiedStyle(Button cell, String val) {
        cell.getStyleClass()
                .removeAll(
                        "pm-machine-calendar-cell-available",
                        "pm-machine-calendar-cell-occupied");
        if (MachineCalendarCellValues.isOccupied(val)) {
            cell.getStyleClass().add("pm-machine-calendar-cell-occupied");
        } else {
            cell.getStyleClass().add("pm-machine-calendar-cell-available");
        }
    }

    private static String cellKey(CellCoord coord) {
        return coord.slotIso() + "\t" + coord.equipmentKey();
    }

    private static CellCoord coordFromNode(Node node) {
        Node n = node;
        while (n != null) {
            if (n.getUserData() instanceof CellCoord coord) {
                return coord;
            }
            n = n.getParent();
        }
        return null;
    }

    private static String formatSlotLabel(String slotIso) {
        try {
            return LocalDateTime.parse(slotIso).format(SLOT_LABEL);
        } catch (Exception e) {
            return slotIso;
        }
    }

    private void notifyDirty() {
        if (dirtyListener != null) {
            dirtyListener.accept(hasUnsavedEdits());
        }
    }

    private static Map<String, Map<String, String>> deepCopyCells(
            Map<String, Map<String, String>> src) {
        Map<String, Map<String, String>> out = new HashMap<>();
        for (Map.Entry<String, Map<String, String>> e : src.entrySet()) {
            out.put(e.getKey(), new HashMap<>(e.getValue()));
        }
        return out;
    }

    private static Map<String, Object> deepCopyPatch(Map<String, Object> patch) {
        Map<String, Object> out = new HashMap<>();
        out.put("date", patch.get("date"));
        Object rowsObj = patch.get("rows");
        if (rowsObj instanceof List<?> list) {
            List<Map<String, Object>> copy = new ArrayList<>();
            for (Object item : list) {
                if (item instanceof Map<?, ?> row) {
                    Map<String, Object> rowCopy = new HashMap<>();
                    rowCopy.put("slot", row.get("slot"));
                    Object cellsObj = row.get("cells");
                    if (cellsObj instanceof Map<?, ?> cellMap) {
                        Map<String, String> cellCopy = new HashMap<>();
                        for (Map.Entry<?, ?> e : cellMap.entrySet()) {
                            cellCopy.put(String.valueOf(e.getKey()), String.valueOf(e.getValue()));
                        }
                        rowCopy.put("cells", cellCopy);
                    }
                    Object commentsObj = row.get("comments");
                    if (commentsObj instanceof Map<?, ?> commentMap) {
                        Map<String, String> commentCopy = new HashMap<>();
                        for (Map.Entry<?, ?> e : commentMap.entrySet()) {
                            commentCopy.put(
                                    String.valueOf(e.getKey()), String.valueOf(e.getValue()));
                        }
                        rowCopy.put("comments", commentCopy);
                    }
                    copy.add(rowCopy);
                }
            }
            out.put("rows", copy);
        }
        return out;
    }

    private static Label legendChip(String text, String styleClass) {
        Label chip = new Label(text);
        chip.getStyleClass().add("pm-attendance-legend-chip");
        chip.getStyleClass().add(styleClass);
        return chip;
    }
}
