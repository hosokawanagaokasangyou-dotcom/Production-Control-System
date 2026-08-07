package jp.co.pm.ai.desktop.ui;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

import com.fasterxml.jackson.databind.JsonNode;

import javafx.geometry.HPos;
import javafx.geometry.Pos;
import javafx.geometry.VPos;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.layout.ColumnConstraints;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;

/** 1日分の機械カレンダー（時刻行×設備列）。 */
public final class EditableMachineCalendarGridPane extends VBox {

    private static final DateTimeFormatter SLOT_LABEL =
            DateTimeFormatter.ofPattern("HH:mm");

    private final GridPane grid = new GridPane();
    private final ScrollPane scroll = new ScrollPane(grid);

    private LocalDate day;
    private List<ColumnDef> columns = List.of();
    private List<RowDef> rows = new ArrayList<>();
    private Map<String, Map<String, String>> cells = new HashMap<>();
    private int cellSizePx = AttendanceGridCellSizing.DEFAULT_CELL_PX;
    private Consumer<Boolean> dirtyListener;
    private Map<String, Object> savedBaseline = Map.of();

    public record ColumnDef(String equipmentKey, String process, String machine) {}

    public record RowDef(String slotIso) {}

    public EditableMachineCalendarGridPane() {
        getStyleClass().add("pm-machine-calendar-grid");
        setSpacing(6);
        Label legend =
                new Label("クリック: 空 ↔ *（占有）。保存で machine-calendar-data.json に反映。");
        legend.getStyleClass().add("pm-machine-calendar-grid-legend");
        legend.setWrapText(true);
        scroll.setFitToHeight(false);
        scroll.setPrefHeight(480);
        grid.setHgap(2);
        grid.setVgap(2);
        getChildren().addAll(legend, scroll);
    }

    public void setDirtyListener(Consumer<Boolean> listener) {
        this.dirtyListener = listener;
    }

    public void setCellSizePx(int px) {
        int clamped = AttendanceGridCellSizing.clamp(px);
        if (cellSizePx == clamped) {
            return;
        }
        cellSizePx = clamped;
        rebuildGrid();
    }

    public boolean hasUnsavedEdits() {
        return !exportPatchJson().equals(savedBaseline);
    }

    public void captureSavedBaseline() {
        savedBaseline = deepCopyPatch(exportPatchJson());
        notifyDirty();
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
            Map<String, String> outCells = new HashMap<>();
            for (ColumnDef col : columns) {
                String v = rowCells.get(col.equipmentKey());
                if (v != null && !v.isBlank()) {
                    outCells.put(col.equipmentKey(), v);
                }
            }
            rowMap.put("cells", outCells);
            outRows.add(rowMap);
        }
        patch.put("rows", outRows);
        return patch;
    }

    private void rebuildGrid() {
        grid.getChildren().clear();
        grid.getColumnConstraints().clear();
        if (columns.isEmpty() || rows.isEmpty()) {
            return;
        }
        ColumnConstraints timeCol = new ColumnConstraints();
        timeCol.setMinWidth(52);
        timeCol.setPrefWidth(52);
        grid.getColumnConstraints().add(timeCol);
        for (ColumnDef col : columns) {
            ColumnConstraints cc = new ColumnConstraints();
            int w = AttendanceGridCellSizing.memberDayColumnWidth(cellSizePx) + 8;
            cc.setMinWidth(w);
            cc.setPrefWidth(w);
            grid.getColumnConstraints().add(cc);
        }
        Label timeHeader = new Label("時刻");
        timeHeader.getStyleClass().add("pm-machine-calendar-grid-header");
        AttendanceGridCellSizing.applyHeaderLabel(timeHeader, cellSizePx);
        grid.add(timeHeader, 0, 0);
        for (int c = 0; c < columns.size(); c++) {
            ColumnDef col = columns.get(c);
            String header = col.machine().isBlank() ? col.equipmentKey() : col.machine();
            Label h = new Label(header);
            h.getStyleClass().add("pm-machine-calendar-grid-header");
            AttendanceGridCellSizing.applyHeaderLabel(h, cellSizePx);
            h.setWrapText(true);
            GridPane.setHalignment(h, HPos.CENTER);
            grid.add(h, c + 1, 0);
        }
        for (int r = 0; r < rows.size(); r++) {
            RowDef row = rows.get(r);
            Label timeLabel = new Label(formatSlotLabel(row.slotIso()));
            timeLabel.getStyleClass().add("pm-machine-calendar-grid-time");
            AttendanceGridCellSizing.applyMemberNameLabel(timeLabel, cellSizePx);
            grid.add(timeLabel, 0, r + 1);
            Map<String, String> rowCells =
                    cells.computeIfAbsent(row.slotIso(), k -> new HashMap<>());
            for (int c = 0; c < columns.size(); c++) {
                ColumnDef col = columns.get(c);
                String val = rowCells.getOrDefault(col.equipmentKey(), "");
                Button cell = new Button(shortLabel(val));
                cell.getStyleClass().add("pm-machine-calendar-cell");
                AttendanceGridCellSizing.applyMemberCell(cell, cellSizePx);
                if (!val.isBlank()) {
                    cell.getStyleClass().add("pm-machine-calendar-cell-occupied");
                }
                cell.setOnAction(
                        e -> {
                            String next = toggleValue(val);
                            if (next.isEmpty()) {
                                rowCells.remove(col.equipmentKey());
                            } else {
                                rowCells.put(col.equipmentKey(), next);
                            }
                            rebuildGrid();
                            notifyDirty();
                        });
                grid.add(cell, c + 1, r + 1);
            }
        }
    }

    private static String formatSlotLabel(String slotIso) {
        try {
            return LocalDateTime.parse(slotIso).format(SLOT_LABEL);
        } catch (Exception e) {
            return slotIso;
        }
    }

    private static String shortLabel(String val) {
        if (val == null || val.isBlank()) {
            return "·";
        }
        return val.length() > 3 ? val.substring(0, 3) : val;
    }

    private static String toggleValue(String current) {
        if (current == null || current.isBlank()) {
            return "*";
        }
        return "";
    }

    private void notifyDirty() {
        if (dirtyListener != null) {
            dirtyListener.accept(hasUnsavedEdits());
        }
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
                    copy.add(rowCopy);
                }
            }
            out.put("rows", copy);
        }
        return out;
    }
}
