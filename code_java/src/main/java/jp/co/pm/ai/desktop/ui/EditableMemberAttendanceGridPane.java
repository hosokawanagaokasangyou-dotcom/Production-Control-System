package jp.co.pm.ai.desktop.ui;

import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.function.Consumer;

import com.fasterxml.jackson.databind.JsonNode;

import javafx.animation.PauseTransition;
import javafx.geometry.HPos;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.layout.ColumnConstraints;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.util.Duration;

/** メンバー×日付の勤怠プリセット編集グリッド。 */
public final class EditableMemberAttendanceGridPane extends VBox {

    public static final String PRESET_WORK = "WORK";
    public static final String PRESET_OFF_FULL = "OFF_FULL";
    public static final String PRESET_OFF_AM = "OFF_AM";
    public static final String PRESET_OFF_PM = "OFF_PM";
    public static final String PRESET_NO_DISPATCH = "NO_DISPATCH";

    public static final String KIND_PUBLIC = "public_holiday";
    public static final String KIND_SPECIAL = "special_holiday";

    private static final String[] PRESET_CYCLE =
            new String[] {
                PRESET_WORK,
                PRESET_OFF_FULL,
                PRESET_OFF_AM,
                PRESET_OFF_PM,
                PRESET_NO_DISPATCH
            };

    private final GridPane grid = new GridPane();
    private final ScrollPane scroll = new ScrollPane(grid);
    private final StackPane scrollHost = new StackPane();
    private final Region loadingOverlay = new Region();
    private final Label loadingLabel = new Label("読込中…");
    private final StackPane loadingOverlayStack = new StackPane(loadingOverlay, loadingLabel);
    private final Map<String, Button> cellButtons = new HashMap<>();
    private final PauseTransition singleClickDelay = new PauseTransition(Duration.millis(280));
    private LocalDate pendingClickDate;
    private String pendingClickMember;

    private int year;
    private int month;
    private List<String> members = List.of();
    private List<LocalDate> dates = List.of();
    private final Map<String, Map<String, CellState>> cells = new HashMap<>();
    private Consumer<CellEditRequest> cellDetailHandler;
    private int cellSizePx = AttendanceGridCellSizing.DEFAULT_CELL_PX;
    private Consumer<Boolean> dirtyListener;

    public EditableMemberAttendanceGridPane() {
        getStyleClass().add("pm-member-attendance-grid");
        setSpacing(6);

        Label legend =
                new Label(
                        "クリックで切替: 通常 → 全休 → 前休 → 後休 → 配台外（ダブルクリックで時間別編集）");
        legend.getStyleClass().add("pm-member-attendance-grid-legend");
        legend.setWrapText(true);

        singleClickDelay.setOnFinished(
                e -> {
                    if (pendingClickDate != null && pendingClickMember != null) {
                        cyclePreset(pendingClickDate, pendingClickMember);
                        pendingClickDate = null;
                        pendingClickMember = null;
                    }
                });

        scroll.setFitToHeight(false);
        scroll.setPrefHeight(480);
        grid.setHgap(2);
        grid.setVgap(2);

        loadingOverlay.getStyleClass().add("pm-member-attendance-grid-loading-overlay");
        loadingOverlay.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
        loadingLabel.getStyleClass().add("pm-member-attendance-grid-loading-label");
        loadingOverlayStack.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
        loadingOverlayStack.setVisible(false);
        loadingOverlayStack.setMouseTransparent(true);

        scrollHost.getChildren().addAll(scroll, loadingOverlayStack);

        getChildren().addAll(legend, scrollHost);
    }

    /** 月変更などのグリッド再読込中に暗転オーバーレイを表示する。 */
    public void setGridLoading(boolean loading) {
        loadingOverlayStack.setVisible(loading);
        loadingOverlayStack.setMouseTransparent(!loading);
        scroll.setDisable(loading);
        if (loading) {
            if (!getStyleClass().contains("pm-member-attendance-grid-loading")) {
                getStyleClass().add("pm-member-attendance-grid-loading");
            }
        } else {
            getStyleClass().remove("pm-member-attendance-grid-loading");
        }
    }

    public boolean isGridLoading() {
        return loadingOverlayStack.isVisible();
    }

    public void setCellDetailHandler(Consumer<CellEditRequest> handler) {
        this.cellDetailHandler = handler;
    }

    public void setDirtyListener(Consumer<Boolean> listener) {
        this.dirtyListener = listener;
    }

    public boolean hasUnsavedEdits() {
        for (Map<String, CellState> row : cells.values()) {
            for (CellState st : row.values()) {
                if (st != null && st.manualEdit) {
                    return true;
                }
            }
        }
        return false;
    }

    public void clearUnsavedEditFlags() {
        for (Map<String, CellState> row : cells.values()) {
            for (Map.Entry<String, CellState> e : row.entrySet()) {
                CellState st = e.getValue();
                if (st != null && st.manualEdit) {
                    e.setValue(
                            new CellState(
                                    st.dayPreset,
                                    st.leaveType,
                                    st.companyKind,
                                    false,
                                    st.hourly));
                }
            }
        }
        notifyDirtyChanged();
    }

    private void notifyDirtyChanged() {
        if (dirtyListener != null) {
            dirtyListener.accept(hasUnsavedEdits());
        }
    }

    public int cellSizePx() {
        return cellSizePx;
    }

    public void setCellSizePx(int px) {
        int clamped = AttendanceGridCellSizing.clamp(px);
        if (cellSizePx == clamped) {
            return;
        }
        cellSizePx = clamped;
        rebuildGrid();
    }

    public record CellEditRequest(LocalDate date, String member, CellState state) {}

    public void loadFromMemberGridJson(JsonNode node) {
        year = node.path("year").asInt(LocalDate.now().getYear());
        month = node.path("month").asInt(LocalDate.now().getMonthValue());
        members = new ArrayList<>();
        node.path("members").forEach(m -> members.add(m.asText("")));
        dates = new ArrayList<>();
        node.path("dates").forEach(d -> {
            try {
                dates.add(LocalDate.parse(d.asText()));
            } catch (Exception ignored) {
                // skip
            }
        });
        cells.clear();
        JsonNode cellsNode = node.path("cells");
        cellsNode
                .fields()
                .forEachRemaining(
                        dayEntry -> {
                            String dKey = dayEntry.getKey();
                            JsonNode perMember = dayEntry.getValue();
                            Map<String, CellState> row = new HashMap<>();
                            perMember
                                    .fields()
                                    .forEachRemaining(
                                            me -> {
                                                JsonNode c = me.getValue();
                                                Map<String, String> hourly = new HashMap<>();
                                                JsonNode hourlyNode = c.path("hourly");
                                                if (hourlyNode.isObject()) {
                                                    hourlyNode
                                                            .fields()
                                                            .forEachRemaining(
                                                                    h ->
                                                                            hourly.put(
                                                                                    h.getKey(),
                                                                                    h.getValue()
                                                                                            .asText(
                                                                                                    "")));
                                                }
                                                row.put(
                                                        me.getKey(),
                                                        new CellState(
                                                                c.path("day_preset")
                                                                        .asText(PRESET_WORK),
                                                                c.path("leave_type").asText(""),
                                                                c.path("company_kind").asText(""),
                                                                c.path("manual_edit")
                                                                        .asBoolean(false),
                                                                hourly));
                                            });
                            cells.put(dKey, row);
                        });
        rebuildGrid();
        notifyDirtyChanged();
    }

    public Map<String, Object> exportPatchJson() {
        Map<String, Object> patch = new HashMap<>();
        patch.put("year", year);
        patch.put("month", month);
        Map<String, Object> outCells = new HashMap<>();
        for (LocalDate d : dates) {
            String dKey = d.toString();
            Map<String, CellState> perMember = cells.get(dKey);
            if (perMember == null) {
                continue;
            }
            Map<String, Object> dayMap = new HashMap<>();
            for (String member : members) {
                CellState st = perMember.get(member);
                if (st == null || !st.manualEdit) {
                    continue;
                }
                Map<String, Object> cell = new HashMap<>();
                cell.put("day_preset", st.dayPreset);
                if (!st.hourly.isEmpty()) {
                    cell.put("hourly", new HashMap<>(st.hourly));
                }
                dayMap.put(member, cell);
            }
            if (!dayMap.isEmpty()) {
                outCells.put(dKey, dayMap);
            }
        }
        patch.put("cells", outCells);
        return patch;
    }

    private void rebuildGrid() {
        grid.getChildren().clear();
        grid.getColumnConstraints().clear();
        cellButtons.clear();
        if (members.isEmpty() || dates.isEmpty()) {
            return;
        }

        ColumnConstraints nameCol = new ColumnConstraints();
        nameCol.setMinWidth(AttendanceGridCellSizing.memberNameColumnWidth(cellSizePx));
        nameCol.setPrefWidth(AttendanceGridCellSizing.memberNameColumnWidth(cellSizePx));
        grid.getColumnConstraints().add(nameCol);
        for (int i = 0; i < dates.size(); i++) {
            ColumnConstraints cc = new ColumnConstraints();
            int dayW = AttendanceGridCellSizing.memberDayColumnWidth(cellSizePx);
            cc.setMinWidth(dayW);
            cc.setPrefWidth(dayW);
            grid.getColumnConstraints().add(cc);
        }

        Label nameHeader = new Label("メンバー");
        nameHeader.getStyleClass().add("pm-member-attendance-grid-header");
        AttendanceGridCellSizing.applyHeaderLabel(nameHeader, cellSizePx);
        grid.add(nameHeader, 0, 0);
        Locale locale = Locale.JAPAN;
        for (int col = 0; col < dates.size(); col++) {
            LocalDate d = dates.get(col);
            String dow = d.getDayOfWeek().getDisplayName(TextStyle.SHORT, locale);
            Label h = new Label(d.getDayOfMonth() + dow);
            h.getStyleClass().add("pm-member-attendance-grid-header");
            AttendanceGridCellSizing.applyHeaderLabel(h, cellSizePx);
            if (isCompanyOffDay(d)) {
                h.getStyleClass().add("pm-member-attendance-company-off");
            }
            GridPane.setHalignment(h, HPos.CENTER);
            grid.add(h, col + 1, 0);
        }

        for (int row = 0; row < members.size(); row++) {
            String member = members.get(row);
            Label name = new Label(member);
            name.getStyleClass().add("pm-member-attendance-grid-member");
            grid.add(name, 0, row + 1);
            for (int col = 0; col < dates.size(); col++) {
                LocalDate d = dates.get(col);
                String dKey = d.toString();
                CellState st =
                        cells.computeIfAbsent(dKey, k -> new HashMap<>())
                                .computeIfAbsent(
                                        member,
                                        m ->
                                                new CellState(
                                                        PRESET_WORK,
                                                        "通常",
                                                        companyKindFor(d, dKey),
                                                        false,
                                                        Map.of()));
                Button cell = new Button(shortLabel(st));
                cell.getStyleClass().add("pm-member-att-cell");
                AttendanceGridCellSizing.applyMemberCell(cell, cellSizePx);
                applyCellStyle(cell, st);
                if (!st.hourly.isEmpty()) {
                    cell.getStyleClass().add("pm-member-att-cell-hourly");
                }
                cell.setOnMouseClicked(
                        e -> {
                            if (e.getClickCount() >= 2) {
                                singleClickDelay.stop();
                                pendingClickDate = null;
                                pendingClickMember = null;
                                if (cellDetailHandler != null) {
                                    CellState current = cells.get(dKey).get(member);
                                    cellDetailHandler.accept(
                                            new CellEditRequest(d, member, current));
                                }
                                return;
                            }
                            if (e.getClickCount() == 1) {
                                pendingClickDate = d;
                                pendingClickMember = member;
                                singleClickDelay.playFromStart();
                            }
                        });
                cellButtons.put(cellKey(dKey, member), cell);
                grid.add(cell, col + 1, row + 1);
            }
        }
    }

    private void cyclePreset(LocalDate date, String member) {
        String dKey = date.toString();
        CellState prev = cells.get(dKey).get(member);
        String current = prev != null ? prev.dayPreset : PRESET_WORK;
        String next = nextPreset(current);
        Map<String, String> hourly =
                prev != null ? new HashMap<>(prev.hourly) : new HashMap<>();
        String companyKind =
                prev != null && !prev.companyKind.isBlank()
                        ? prev.companyKind
                        : companyKindFor(date, dKey);
        cells.get(dKey)
                .put(
                        member,
                        new CellState(
                                next,
                                defaultLeaveForPreset(next),
                                companyKind,
                                true,
                                hourly));
        updateCellButton(dKey, member);
        notifyDirtyChanged();
    }

    private static String nextPreset(String current) {
        for (int i = 0; i < PRESET_CYCLE.length; i++) {
            if (PRESET_CYCLE[i].equals(current)) {
                return PRESET_CYCLE[(i + 1) % PRESET_CYCLE.length];
            }
        }
        return PRESET_OFF_FULL;
    }

    private void updateCellButton(String dKey, String member) {
        Button cell = cellButtons.get(cellKey(dKey, member));
        CellState st = cells.get(dKey).get(member);
        if (cell == null || st == null) {
            return;
        }
        cell.setText(shortLabel(st));
        applyCellStyle(cell, st);
        cell.getStyleClass().remove("pm-member-att-cell-hourly");
        if (!st.hourly.isEmpty()) {
            cell.getStyleClass().add("pm-member-att-cell-hourly");
        }
    }

    private static String cellKey(String dKey, String member) {
        return dKey + "\t" + member;
    }

    private boolean isCompanyOffDay(LocalDate d) {
        String dKey = d.toString();
        for (String m : members) {
            Map<String, CellState> row = cells.get(dKey);
            if (row != null && row.get(m) != null) {
                String k = row.get(m).companyKind;
                return KIND_PUBLIC.equals(k) || KIND_SPECIAL.equals(k);
            }
        }
        return d.getDayOfWeek().getValue() >= 6;
    }

    private String companyKindFor(LocalDate d, String dKey) {
        for (String m : members) {
            Map<String, CellState> row = cells.get(dKey);
            if (row != null && row.get(m) != null && !row.get(m).companyKind.isBlank()) {
                return row.get(m).companyKind;
            }
        }
        if (d.getDayOfWeek().getValue() >= 6) {
            return KIND_PUBLIC;
        }
        return "working_day";
    }

    private static String defaultLeaveForPreset(String preset) {
        return switch (preset) {
            case PRESET_OFF_FULL -> "休";
            case PRESET_OFF_AM -> "前休";
            case PRESET_OFF_PM -> "後休";
            case PRESET_NO_DISPATCH -> "-";
            default -> "通常";
        };
    }

    private static String shortLabel(CellState st) {
        if (PRESET_OFF_FULL.equals(st.dayPreset)) {
            return "休";
        }
        if (PRESET_OFF_AM.equals(st.dayPreset)) {
            return "前";
        }
        if (PRESET_OFF_PM.equals(st.dayPreset)) {
            return "後";
        }
        if (PRESET_NO_DISPATCH.equals(st.dayPreset)) {
            return "-";
        }
        if (!st.hourly.isEmpty()) {
            return "時";
        }
        return "·";
    }

    private static void applyCellStyle(Button cell, CellState st) {
        cell.getStyleClass()
                .removeAll(
                        "pm-member-att-cell-work",
                        "pm-member-att-cell-off",
                        "pm-member-att-cell-partial",
                        "pm-member-att-cell-nodispatch");
        switch (st.dayPreset) {
            case PRESET_OFF_FULL -> cell.getStyleClass().add("pm-member-att-cell-off");
            case PRESET_OFF_AM, PRESET_OFF_PM -> cell.getStyleClass().add("pm-member-att-cell-partial");
            case PRESET_NO_DISPATCH -> cell.getStyleClass().add("pm-member-att-cell-nodispatch");
            default -> cell.getStyleClass().add("pm-member-att-cell-work");
        }
    }

    public void applyHourlyEdit(
            LocalDate date, String member, Map<String, String> hourly, String dayPreset) {
        String dKey = date.toString();
        CellState prev =
                cells.computeIfAbsent(dKey, k -> new HashMap<>())
                        .getOrDefault(
                                member,
                                new CellState(
                                        PRESET_WORK,
                                        "通常",
                                        companyKindFor(date, dKey),
                                        false,
                                        Map.of()));
        cells.get(dKey)
                .put(
                        member,
                        new CellState(
                                dayPreset != null ? dayPreset : prev.dayPreset,
                                prev.leaveType,
                                prev.companyKind,
                                true,
                                new HashMap<>(hourly)));
        updateCellButton(dKey, member);
        notifyDirtyChanged();
    }

    public record CellState(
            String dayPreset,
            String leaveType,
            String companyKind,
            boolean manualEdit,
            Map<String, String> hourly) {}
}
