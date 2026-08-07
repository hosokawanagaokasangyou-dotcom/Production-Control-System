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
import javafx.geometry.VPos;
import javafx.scene.Node;
import javafx.scene.control.Button;
import javafx.scene.control.ContextMenu;
import javafx.scene.control.Label;
import javafx.scene.control.Menu;
import javafx.scene.control.MenuItem;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.SeparatorMenuItem;
import javafx.scene.control.Tooltip;
import javafx.scene.layout.ColumnConstraints;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.scene.shape.Polygon;
import javafx.stage.Stage;
import javafx.stage.Window;
import javafx.util.Duration;

/** メンバー×日付の勤怠プリセット編集グリッド。 */
public final class EditableMemberAttendanceGridPane extends VBox {

    public static final String PRESET_WORK = "WORK";
    public static final String PRESET_OFF_FULL = "OFF_FULL";
    public static final String PRESET_PAID_LEAVE = "PAID_LEAVE";
    public static final String PRESET_ABSENT = "ABSENT";
    public static final String PRESET_OFF_AM = "OFF_AM";
    public static final String PRESET_OFF_PM = "OFF_PM";
    public static final String PRESET_NO_DISPATCH = "NO_DISPATCH";
    public static final String PRESET_HOLIDAY_WORK = "HOLIDAY_WORK";
    public static final String PRESET_HOLIDAY_WORK_AM = "HOLIDAY_WORK_AM";
    public static final String PRESET_HOLIDAY_WORK_PM = "HOLIDAY_WORK_PM";

    public static final String KIND_PUBLIC = "public_holiday";
    public static final String KIND_SPECIAL = "special_holiday";

    private static final String[] PRESET_CYCLE =
            new String[] {
                PRESET_WORK,
                PRESET_OFF_FULL,
                PRESET_PAID_LEAVE,
                PRESET_ABSENT,
                PRESET_OFF_AM,
                PRESET_OFF_PM,
                PRESET_HOLIDAY_WORK,
                PRESET_HOLIDAY_WORK_AM,
                PRESET_HOLIDAY_WORK_PM
            };

    private record PresetMenuOption(String preset, String label) {}

    private static final List<PresetMenuOption> PRESET_MENU_OPTIONS =
            List.of(
                    new PresetMenuOption(PRESET_WORK, "· 通常"),
                    new PresetMenuOption(PRESET_OFF_FULL, "休 全休"),
                    new PresetMenuOption(PRESET_PAID_LEAVE, "年 有給休暇(年休)"),
                    new PresetMenuOption(PRESET_ABSENT, "欠 欠勤"),
                    new PresetMenuOption(PRESET_OFF_AM, "前 前休"),
                    new PresetMenuOption(PRESET_OFF_PM, "後 後休"),
                    new PresetMenuOption(PRESET_HOLIDAY_WORK, "休出 休日出勤"),
                    new PresetMenuOption(PRESET_HOLIDAY_WORK_AM, "前出 午前休出"),
                    new PresetMenuOption(PRESET_HOLIDAY_WORK_PM, "後出 午後休出"));

    private final GridPane grid = new GridPane();
    private final ScrollPane scroll = new ScrollPane(grid);
    private final StackPane scrollHost = new StackPane();
    private final AttendanceGridLoadingOverlay loadingOverlay =
            new AttendanceGridLoadingOverlay("pm-member-attendance-grid-loading-overlay");
    private final Map<String, CellUi> cellUiMap = new HashMap<>();
    private final PauseTransition singleClickDelay = new PauseTransition(Duration.millis(280));
    private LocalDate pendingClickDate;
    private String pendingClickMember;

    private int year;
    private int month;
    private List<String> members = List.of();
    private Map<String, String> primaryRoles = new HashMap<>();
    private String selectedMember = null;
    private List<LocalDate> dates = List.of();
    private final Map<String, Map<String, CellState>> cells = new HashMap<>();
    /** コメントを明示編集したセル（空文字での削除を保存に反映するため）。 */
    private final java.util.Set<String> commentEditedCellKeys = new java.util.HashSet<>();
    /** 時間別を明示編集／クリアしたセル（空 map での削除を保存に反映するため）。 */
    private final java.util.Set<String> hourlyEditedCellKeys = new java.util.HashSet<>();
    private Consumer<CellEditRequest> cellDetailHandler;
    private int cellSizePx = AttendanceGridCellSizing.DEFAULT_CELL_PX;
    private Consumer<Boolean> dirtyListener;
    /** 読込／保存直後の exportPatchJson スナップショット（JSON 正本との差分で未保存を判定）。 */
    private Map<String, Object> savedBaselinePatch = Map.of();
    private Window commentDialogOwner;
    private final GridRowHoverDimmingController rowDimming = new GridRowHoverDimmingController();

    private record CellUi(Button button, Polygon commentMark) {}

    public EditableMemberAttendanceGridPane() {
        getStyleClass().add("pm-member-attendance-grid");
        setSpacing(6);

        HBox legendChips =
                new HBox(
                        6,
                        legendChip("· 通常", "pm-att-legend-work"),
                        legendChip("休", "pm-att-legend-off"),
                        legendChip("年休", "pm-att-legend-paid-leave"),
                        legendChip("欠勤", "pm-att-legend-absent"),
                        legendChip("前/後", "pm-att-legend-partial"),
                        legendChip("休出", "pm-att-legend-holiday-work"),
                        legendChip("前出/後出", "pm-att-legend-holiday-work-partial"));
        legendChips.getStyleClass().add("pm-attendance-legend-chips");
        Label legend =
                new Label(
                        "クリック: ·→休→年休→欠→前→後→休出→前出→後出　｜ダブルクリック: 時間別（青枠=時間別あり）　｜右クリック: コメント（▲=あり）");
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

        scrollHost.getChildren().addAll(scroll, loadingOverlay);
        rowDimming.installScrollClearOnExit(scroll);

        getChildren().addAll(legendChips, legend, scrollHost);
    }

    public void refreshRowHoverDimming() {
        rowDimming.refresh();
    }

    /** 月変更などのグリッド再読込中に暗転オーバーレイを表示する。 */
    public void setGridLoading(boolean loading) {
        setGridLoading(loading, null);
    }

    public void setGridLoading(boolean loading, String message) {
        loadingOverlay.setLoading(loading, message);
        scroll.setDisable(loading);
        if (loading) {
            rowDimming.setHoveredRow(-1);
            if (!getStyleClass().contains("pm-member-attendance-grid-loading")) {
                getStyleClass().add("pm-member-attendance-grid-loading");
            }
        } else {
            getStyleClass().remove("pm-member-attendance-grid-loading");
        }
    }

    public void setGridLoadingMessage(String message) {
        if (loadingOverlay.isVisible()) {
            loadingOverlay.setMessage(message);
        }
    }

    public boolean isGridLoading() {
        return loadingOverlay.isVisible();
    }

    public void setCellDetailHandler(Consumer<CellEditRequest> handler) {
        this.cellDetailHandler = handler;
    }

    public void setDirtyListener(Consumer<Boolean> listener) {
        this.dirtyListener = listener;
    }

    public void setCommentDialogOwner(Window owner) {
        this.commentDialogOwner = owner;
    }

    public boolean hasUnsavedEdits() {
        return !patchEquals(exportPatchJson(), savedBaselinePatch);
    }

    public void clearUnsavedEditFlags() {
        captureSavedBaseline();
        notifyDirtyChanged();
    }

    private void captureSavedBaseline() {
        savedBaselinePatch = deepCopyPatch(exportPatchJson());
    }

    private static Map<String, Object> deepCopyPatch(Map<String, Object> patch) {
        Map<String, Object> out = new HashMap<>();
        out.put("year", patch.get("year"));
        out.put("month", patch.get("month"));
        Object cellsObj = patch.get("cells");
        if (cellsObj instanceof Map<?, ?> cells) {
            Map<String, Object> outCells = new HashMap<>();
            for (Map.Entry<?, ?> dayEntry : cells.entrySet()) {
                String dKey = String.valueOf(dayEntry.getKey());
                Map<String, Object> outDay = new HashMap<>();
                if (dayEntry.getValue() instanceof Map<?, ?> dayMap) {
                    for (Map.Entry<?, ?> memberEntry : dayMap.entrySet()) {
                        String member = String.valueOf(memberEntry.getKey());
                        if (memberEntry.getValue() instanceof Map<?, ?> cellMap) {
                            Map<String, Object> outCell = new HashMap<>();
                            for (Map.Entry<?, ?> field : cellMap.entrySet()) {
                                String key = String.valueOf(field.getKey());
                                Object val = field.getValue();
                                if ("hourly".equals(key) && val instanceof Map<?, ?> hourlyMap) {
                                    Map<String, String> hourlyCopy = new HashMap<>();
                                    for (Map.Entry<?, ?> h : hourlyMap.entrySet()) {
                                        hourlyCopy.put(
                                                String.valueOf(h.getKey()),
                                                String.valueOf(h.getValue()));
                                    }
                                    outCell.put(key, hourlyCopy);
                                } else {
                                    outCell.put(key, val);
                                }
                            }
                            outDay.put(member, outCell);
                        }
                    }
                }
                outCells.put(dKey, outDay);
            }
            out.put("cells", outCells);
        } else {
            out.put("cells", new HashMap<>());
        }
        Object rosterObj = patch.get("member_roster");
        if (rosterObj instanceof List<?> roster) {
            List<Map<String, Object>> rosterCopy = new ArrayList<>();
            for (Object item : roster) {
                if (item instanceof Map<?, ?> ent) {
                    Map<String, Object> row = new HashMap<>();
                    row.put("name", String.valueOf(ent.get("name")));
                    row.put(
                            "primary_role",
                            String.valueOf(
                                    ent.get("primary_role") != null
                                            ? ent.get("primary_role")
                                            : MemberAttendanceMemberEditDialog.ROLE_POST));
                    rosterCopy.add(row);
                }
            }
            out.put("member_roster", rosterCopy);
        } else {
            out.put("member_roster", List.of());
        }
        return out;
    }

    private static boolean patchEquals(Map<String, Object> a, Map<String, Object> b) {
        if (patchInt(a, "year") != patchInt(b, "year")) {
            return false;
        }
        if (patchInt(a, "month") != patchInt(b, "month")) {
            return false;
        }
        if (!patchRosterEqual(a.get("member_roster"), b.get("member_roster"))) {
            return false;
        }
        return patchCellMapsEqual(patchCells(a), patchCells(b));
    }

    private static boolean patchRosterEqual(Object a, Object b) {
        List<Map<String, String>> la = rosterToList(a);
        List<Map<String, String>> lb = rosterToList(b);
        if (la.size() != lb.size()) {
            return false;
        }
        for (int i = 0; i < la.size(); i++) {
            Map<String, String> ea = la.get(i);
            Map<String, String> eb = lb.get(i);
            if (!nz(ea.get("name")).equals(nz(eb.get("name")))
                    || !nz(ea.get("primary_role")).equals(nz(eb.get("primary_role")))) {
                return false;
            }
        }
        return true;
    }

    private static List<Map<String, String>> rosterToList(Object rosterObj) {
        List<Map<String, String>> out = new ArrayList<>();
        if (rosterObj instanceof List<?> roster) {
            for (Object item : roster) {
                if (item instanceof Map<?, ?> ent) {
                    Map<String, String> row = new HashMap<>();
                    row.put("name", nz(ent.get("name")));
                    row.put("primary_role", nz(ent.get("primary_role")));
                    out.add(row);
                }
            }
        }
        return out;
    }

    private static int patchInt(Map<String, Object> patch, String key) {
        Object v = patch.get(key);
        if (v instanceof Number n) {
            return n.intValue();
        }
        try {
            return Integer.parseInt(String.valueOf(v));
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    @SuppressWarnings("unchecked")
    private static Map<String, Object> patchCells(Map<String, Object> patch) {
        Object cells = patch.get("cells");
        if (cells instanceof Map<?, ?> m) {
            return (Map<String, Object>) m;
        }
        return Map.of();
    }

    private static boolean patchCellMapsEqual(
            Map<String, Object> a, Map<String, Object> b) {
        if (a.size() != b.size()) {
            return false;
        }
        for (Map.Entry<String, Object> dayEntry : a.entrySet()) {
            Object rowB = b.get(dayEntry.getKey());
            if (!(rowB instanceof Map<?, ?> memberMapB)) {
                return false;
            }
            if (!(dayEntry.getValue() instanceof Map<?, ?> memberMapA)) {
                return false;
            }
            if (memberMapA.size() != memberMapB.size()) {
                return false;
            }
            for (Map.Entry<?, ?> memberEntry : memberMapA.entrySet()) {
                String member = String.valueOf(memberEntry.getKey());
                Object cellB = memberMapB.get(member);
                if (!(memberEntry.getValue() instanceof Map<?, ?> cellA)
                        || !(cellB instanceof Map<?, ?> cellBMap)) {
                    return false;
                }
                if (!patchCellEquals(cellA, cellBMap)) {
                    return false;
                }
            }
        }
        return true;
    }

    private static boolean patchCellEquals(Map<?, ?> a, Map<?, ?> b) {
        if (!nz(a.get("day_preset")).equals(nz(b.get("day_preset")))) {
            return false;
        }
        if (!nz(a.get("comment")).equals(nz(b.get("comment")))) {
            return false;
        }
        return hourlyMapsEqual(a.get("hourly"), b.get("hourly"));
    }

    private static boolean hourlyMapsEqual(Object a, Object b) {
        Map<String, String> ma = hourlyToStringMap(a);
        Map<String, String> mb = hourlyToStringMap(b);
        if (ma.size() != mb.size()) {
            return false;
        }
        for (Map.Entry<String, String> e : ma.entrySet()) {
            if (!nz(mb.get(e.getKey())).equals(nz(e.getValue()))) {
                return false;
            }
        }
        return true;
    }

    private static Map<String, String> hourlyToStringMap(Object hourly) {
        Map<String, String> out = new HashMap<>();
        if (hourly instanceof Map<?, ?> m) {
            for (Map.Entry<?, ?> e : m.entrySet()) {
                out.put(String.valueOf(e.getKey()), String.valueOf(e.getValue()));
            }
        }
        return out;
    }

    private static String nz(Object v) {
        return v == null ? "" : String.valueOf(v).trim();
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
        primaryRoles = new HashMap<>();
        node.path("members").forEach(m -> members.add(m.asText("")));
        JsonNode rosterNode = node.path("member_roster");
        if (rosterNode.isArray() && rosterNode.size() > 0) {
            members.clear();
            rosterNode.forEach(
                    ent -> {
                        String name = ent.path("name").asText("").trim();
                        if (name.isEmpty()) {
                            return;
                        }
                        members.add(name);
                        String role =
                                ent.path("primary_role")
                                        .asText(MemberAttendanceMemberEditDialog.ROLE_POST)
                                        .trim();
                        if (!role.equals(MemberAttendanceMemberEditDialog.ROLE_LOGISTICS)) {
                            role = MemberAttendanceMemberEditDialog.ROLE_POST;
                        }
                        primaryRoles.put(name, role);
                    });
        } else {
            JsonNode rolesNode = node.path("primary_roles");
            if (rolesNode.isObject()) {
                rolesNode
                        .fields()
                        .forEachRemaining(
                                e ->
                                        primaryRoles.put(
                                                e.getKey(),
                                                e.getValue()
                                                        .asText(
                                                                MemberAttendanceMemberEditDialog
                                                                        .ROLE_POST)));
            }
            for (String member : members) {
                primaryRoles.putIfAbsent(
                        member, MemberAttendanceMemberEditDialog.ROLE_POST);
            }
        }
        dates = new ArrayList<>();
        node.path("dates").forEach(d -> {
            try {
                dates.add(LocalDate.parse(d.asText()));
            } catch (Exception ignored) {
                // skip
            }
        });
        cells.clear();
        commentEditedCellKeys.clear();
        hourlyEditedCellKeys.clear();
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
                                                                hourly,
                                                                c.path("comment").asText("")));
                                            });
                            cells.put(dKey, row);
                        });
        rebuildGrid();
        captureSavedBaseline();
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
                if (!st.comment.isBlank()) {
                    cell.put("comment", st.comment);
                } else if (commentEditedCellKeys.contains(cellKey(dKey, member))) {
                    cell.put("comment", "");
                }
                if (!st.hourly.isEmpty()) {
                    cell.put("hourly", new HashMap<>(st.hourly));
                } else if (hourlyEditedCellKeys.contains(cellKey(dKey, member))) {
                    cell.put("hourly", new HashMap<>());
                }
                dayMap.put(member, cell);
            }
            if (!dayMap.isEmpty()) {
                outCells.put(dKey, dayMap);
            }
        }
        patch.put("cells", outCells);
        patch.put("member_roster", exportMemberRoster());
        return patch;
    }

    public String primaryRoleFor(String member) {
        if (member == null) {
            return MemberAttendanceMemberEditDialog.ROLE_POST;
        }
        return primaryRoles.getOrDefault(member, MemberAttendanceMemberEditDialog.ROLE_POST);
    }

    private List<Map<String, Object>> exportMemberRoster() {
        List<Map<String, Object>> roster = new ArrayList<>();
        for (String member : members) {
            Map<String, Object> row = new HashMap<>();
            row.put("name", member);
            row.put(
                    "primary_role",
                    primaryRoles.getOrDefault(
                            member, MemberAttendanceMemberEditDialog.ROLE_POST));
            roster.add(row);
        }
        return roster;
    }

    public String selectedMemberName() {
        return selectedMember;
    }

    public void addMember(String name, String primaryRole) {
        String n = name != null ? name.trim() : "";
        if (n.isEmpty() || members.contains(n)) {
            return;
        }
        String role =
                MemberAttendanceMemberEditDialog.ROLE_LOGISTICS.equals(primaryRole)
                        ? MemberAttendanceMemberEditDialog.ROLE_LOGISTICS
                        : MemberAttendanceMemberEditDialog.ROLE_POST;
        members.add(n);
        primaryRoles.put(n, role);
        selectedMember = n;
        rebuildGrid();
        notifyDirtyChanged();
    }

    public void updateMember(String oldName, String newName, String primaryRole) {
        if (oldName == null || newName == null) {
            return;
        }
        String oldN = oldName.trim();
        String newN = newName.trim();
        if (oldN.isEmpty() || newN.isEmpty()) {
            return;
        }
        int idx = members.indexOf(oldN);
        if (idx < 0) {
            return;
        }
        if (!oldN.equals(newN) && members.contains(newN)) {
            return;
        }
        String role =
                MemberAttendanceMemberEditDialog.ROLE_LOGISTICS.equals(primaryRole)
                        ? MemberAttendanceMemberEditDialog.ROLE_LOGISTICS
                        : MemberAttendanceMemberEditDialog.ROLE_POST;
        members.set(idx, newN);
        primaryRoles.remove(oldN);
        primaryRoles.put(newN, role);
        if (!oldN.equals(newN)) {
            migrateMemberCells(oldN, newN);
        }
        if (selectedMember != null && selectedMember.equals(oldN)) {
            selectedMember = newN;
        }
        rebuildGrid();
        notifyDirtyChanged();
    }

    public void removeMember(String name) {
        if (name == null || name.isBlank()) {
            return;
        }
        String n = name.trim();
        if (!members.contains(n)) {
            return;
        }
        members.remove(n);
        primaryRoles.remove(n);
        for (Map<String, CellState> day : cells.values()) {
            day.remove(n);
        }
        if (selectedMember != null && selectedMember.equals(n)) {
            selectedMember = null;
        }
        rebuildGrid();
        notifyDirtyChanged();
    }

    private void migrateMemberCells(String oldName, String newName) {
        for (Map<String, CellState> day : cells.values()) {
            CellState st = day.remove(oldName);
            if (st != null && !day.containsKey(newName)) {
                day.put(newName, st);
            }
        }
        for (String key : new ArrayList<>(commentEditedCellKeys)) {
            if (key.endsWith("|".concat(oldName))) {
                commentEditedCellKeys.remove(key);
                commentEditedCellKeys.add(cellKey(key.substring(0, key.length() - oldName.length() - 1), newName));
            }
        }
        for (String key : new ArrayList<>(hourlyEditedCellKeys)) {
            if (key.endsWith("|".concat(oldName))) {
                hourlyEditedCellKeys.remove(key);
                hourlyEditedCellKeys.add(cellKey(key.substring(0, key.length() - oldName.length() - 1), newName));
            }
        }
    }

    private void selectMember(String member) {
        selectedMember = member;
        refreshMemberSelectionStyles();
    }

    private void refreshMemberSelectionStyles() {
        grid.getChildren().stream()
                .filter(n -> n instanceof Label)
                .map(n -> (Label) n)
                .filter(l -> l.getStyleClass().contains("pm-member-attendance-grid-member"))
                .forEach(
                        l -> {
                            boolean sel =
                                    selectedMember != null
                                            && selectedMember.equals(l.getText());
                            if (sel) {
                                if (!l.getStyleClass()
                                        .contains("pm-member-attendance-grid-member-selected")) {
                                    l.getStyleClass()
                                            .add("pm-member-attendance-grid-member-selected");
                                }
                            } else {
                                l.getStyleClass()
                                        .remove("pm-member-attendance-grid-member-selected");
                            }
                        });
    }

    private void rebuildGrid() {
        grid.getChildren().clear();
        grid.getColumnConstraints().clear();
        cellUiMap.clear();
        rowDimming.clear();
        if (members.isEmpty() || dates.isEmpty()) {
            return;
        }

        ColumnConstraints roleCol = new ColumnConstraints();
        roleCol.setMinWidth(AttendanceGridCellSizing.memberPrimaryRoleColumnWidth(cellSizePx));
        roleCol.setPrefWidth(AttendanceGridCellSizing.memberPrimaryRoleColumnWidth(cellSizePx));
        grid.getColumnConstraints().add(roleCol);
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

        Label roleHeader = new Label("主担当");
        roleHeader.getStyleClass().add("pm-member-attendance-grid-header");
        AttendanceGridCellSizing.applyHeaderLabel(roleHeader, cellSizePx);
        grid.add(roleHeader, 0, 0);
        Label nameHeader = new Label("メンバー");
        nameHeader.getStyleClass().add("pm-member-attendance-grid-header");
        AttendanceGridCellSizing.applyHeaderLabel(nameHeader, cellSizePx);
        grid.add(nameHeader, 1, 0);
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
            grid.add(h, col + 2, 0);
        }

        for (int row = 0; row < members.size(); row++) {
            String member = members.get(row);
            Region rowBand = new Region();
            rowBand.getStyleClass().add(GridRowHoverDimmingController.STYLE_BAND);
            rowBand.setMouseTransparent(true);
            rowBand.setMaxWidth(Double.MAX_VALUE);
            rowBand.setMaxHeight(Double.MAX_VALUE);
            GridPane.setColumnSpan(rowBand, dates.size() + 2);
            grid.add(rowBand, 0, row + 1);

            Label roleLabel =
                    new Label(
                            primaryRoles.getOrDefault(
                                    member, MemberAttendanceMemberEditDialog.ROLE_POST));
            roleLabel.getStyleClass().add("pm-member-attendance-grid-role");
            AttendanceGridCellSizing.applyMemberNameLabel(roleLabel, cellSizePx);
            GridPane.setHalignment(roleLabel, HPos.CENTER);
            GridPane.setValignment(roleLabel, VPos.CENTER);
            grid.add(roleLabel, 0, row + 1);
            rowDimming.installHover(roleLabel, row);

            Label name = new Label(member);
            name.getStyleClass().add("pm-member-attendance-grid-member");
            if (selectedMember != null && selectedMember.equals(member)) {
                name.getStyleClass().add("pm-member-attendance-grid-member-selected");
            }
            AttendanceGridCellSizing.applyMemberNameLabel(name, cellSizePx);
            GridPane.setHalignment(name, HPos.CENTER);
            GridPane.setValignment(name, VPos.CENTER);
            name.setOnMouseClicked(e -> selectMember(member));
            grid.add(name, 1, row + 1);
            List<StackPane> rowWraps = new ArrayList<>();
            rowDimming.installHover(name, row);
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
                                                        Map.of(),
                                                        ""));
                Button cell = new Button(shortLabel(st));
                cell.getStyleClass().add("pm-member-att-cell");
                AttendanceGridCellSizing.applyMemberCell(cell, cellSizePx);
                applyCellStyle(cell, st);
                if (!st.hourly.isEmpty()) {
                    cell.getStyleClass().add("pm-member-att-cell-hourly");
                }
                applyCellTooltip(cell, st);
                Polygon commentMark = buildCommentMark();
                commentMark.setVisible(hasComment(st.comment));
                StackPane cellWrap = new StackPane(cell, commentMark);
                cellWrap.setMaxSize(Region.USE_PREF_SIZE, Region.USE_PREF_SIZE);
                installCellInteractions(cell, d, member, dKey);
                cellUiMap.put(cellKey(dKey, member), new CellUi(cell, commentMark));
                rowDimming.installHover(cellWrap, row);
                rowWraps.add(cellWrap);
                grid.add(cellWrap, col + 2, row + 1);
            }
            rowDimming.addRow(rowBand, name, new ArrayList<>(rowWraps));
        }
    }

    private void cyclePreset(LocalDate date, String member) {
        String dKey = date.toString();
        Map<String, CellState> row = cells.get(dKey);
        CellState prev = row != null ? row.get(member) : null;
        String current = prev != null ? prev.dayPreset : PRESET_WORK;
        applyPreset(date, member, nextPreset(current));
    }

    private void applyPreset(LocalDate date, String member, String preset) {
        if (preset == null || preset.isBlank()) {
            return;
        }
        singleClickDelay.stop();
        pendingClickDate = null;
        pendingClickMember = null;
        String dKey = date.toString();
        CellState prev =
                cells.computeIfAbsent(dKey, k -> new HashMap<>()).get(member);
        if (prev != null && preset.equals(prev.dayPreset)) {
            return;
        }
        Map<String, String> hourly =
                prev != null ? new HashMap<>(prev.hourly) : new HashMap<>();
        String comment = prev != null ? prev.comment : "";
        String companyKind =
                prev != null && !prev.companyKind.isBlank()
                        ? prev.companyKind
                        : companyKindFor(date, dKey);
        cells.get(dKey)
                .put(
                        member,
                        new CellState(
                                preset,
                                defaultLeaveForPreset(preset),
                                companyKind,
                                true,
                                hourly,
                                comment));
        updateCellButton(dKey, member);
        notifyDirtyChanged();
    }

    private static String nextPreset(String current) {
        if (PRESET_NO_DISPATCH.equals(current)) {
            return PRESET_WORK;
        }
        for (int i = 0; i < PRESET_CYCLE.length; i++) {
            if (PRESET_CYCLE[i].equals(current)) {
                return PRESET_CYCLE[(i + 1) % PRESET_CYCLE.length];
            }
        }
        return PRESET_WORK;
    }

    private void updateCellButton(String dKey, String member) {
        CellUi ui = cellUiMap.get(cellKey(dKey, member));
        CellState st = cells.get(dKey).get(member);
        if (ui == null || st == null) {
            return;
        }
        Button cell = ui.button();
        cell.setText(shortLabel(st));
        applyCellStyle(cell, st);
        cell.getStyleClass().remove("pm-member-att-cell-hourly");
        if (!st.hourly.isEmpty()) {
            cell.getStyleClass().add("pm-member-att-cell-hourly");
        }
        applyCellTooltip(cell, st);
        ui.commentMark().setVisible(hasComment(st.comment));
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
            case PRESET_PAID_LEAVE -> "年休";
            case PRESET_ABSENT -> "欠勤";
            case PRESET_OFF_AM -> "前休";
            case PRESET_OFF_PM -> "後休";
            case PRESET_HOLIDAY_WORK -> "休日出勤";
            case PRESET_HOLIDAY_WORK_AM -> "午前休出";
            case PRESET_HOLIDAY_WORK_PM -> "午後休出";
            case PRESET_NO_DISPATCH -> "-";
            default -> "通常";
        };
    }

    private static String shortLabel(CellState st) {
        if (PRESET_OFF_FULL.equals(st.dayPreset)) {
            return "休";
        }
        if (PRESET_PAID_LEAVE.equals(st.dayPreset)) {
            return "年休";
        }
        if (PRESET_ABSENT.equals(st.dayPreset)) {
            return "欠";
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
        if (PRESET_HOLIDAY_WORK.equals(st.dayPreset)) {
            return "休出";
        }
        if (PRESET_HOLIDAY_WORK_AM.equals(st.dayPreset)) {
            return "前出";
        }
        if (PRESET_HOLIDAY_WORK_PM.equals(st.dayPreset)) {
            return "後出";
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
                        "pm-member-att-cell-paid-leave",
                        "pm-member-att-cell-absent",
                        "pm-member-att-cell-partial",
                        "pm-member-att-cell-holiday-work",
                        "pm-member-att-cell-holiday-work-partial",
                        "pm-member-att-cell-nodispatch");
        switch (st.dayPreset) {
            case PRESET_OFF_FULL -> cell.getStyleClass().add("pm-member-att-cell-off");
            case PRESET_PAID_LEAVE -> cell.getStyleClass().add("pm-member-att-cell-paid-leave");
            case PRESET_ABSENT -> cell.getStyleClass().add("pm-member-att-cell-absent");
            case PRESET_OFF_AM, PRESET_OFF_PM -> cell.getStyleClass().add("pm-member-att-cell-partial");
            case PRESET_HOLIDAY_WORK -> cell.getStyleClass().add("pm-member-att-cell-holiday-work");
            case PRESET_HOLIDAY_WORK_AM, PRESET_HOLIDAY_WORK_PM ->
                    cell.getStyleClass().add("pm-member-att-cell-holiday-work-partial");
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
                                        Map.of(),
                                        ""));
        cells.get(dKey)
                .put(
                        member,
                        new CellState(
                                dayPreset != null ? dayPreset : prev.dayPreset,
                                prev.leaveType,
                                prev.companyKind,
                                true,
                                new HashMap<>(hourly),
                                prev.comment));
        hourlyEditedCellKeys.add(cellKey(dKey, member));
        updateCellButton(dKey, member);
        notifyDirtyChanged();
    }

    public void clearHourlyEdit(LocalDate date, String member) {
        applyHourlyEdit(date, member, Map.of(), null);
        CellUi ui = cellUiMap.get(cellKey(date.toString(), member));
        if (ui != null) {
            releaseCellFocus(ui.button());
        }
    }

    public record CellState(
            String dayPreset,
            String leaveType,
            String companyKind,
            boolean manualEdit,
            Map<String, String> hourly,
            String comment) {}

    private static boolean hasComment(String comment) {
        return comment != null && !comment.isBlank();
    }

    private Polygon buildCommentMark() {
        double size = Math.max(7, cellSizePx / 4.5);
        Polygon mark = new Polygon(0, 0, size, 0, size, size);
        mark.getStyleClass().add("pm-member-att-cell-comment-mark");
        mark.setMouseTransparent(true);
        StackPane.setAlignment(mark, Pos.TOP_RIGHT);
        return mark;
    }

    private void installCellInteractions(Button cell, LocalDate date, String member, String dKey) {
        cell.setOnMouseClicked(
                e -> {
                    if (e.getClickCount() >= 2) {
                        singleClickDelay.stop();
                        pendingClickDate = null;
                        pendingClickMember = null;
                        if (cellDetailHandler != null) {
                            CellState current = cells.get(dKey).get(member);
                            cellDetailHandler.accept(new CellEditRequest(date, member, current));
                        }
                        return;
                    }
                    if (e.getClickCount() == 1) {
                        pendingClickDate = date;
                        pendingClickMember = member;
                        singleClickDelay.playFromStart();
                    }
                });
        cell.setOnContextMenuRequested(
                e -> {
                    e.consume();
                    showCellContextMenu(cell, date, member, e.getScreenX(), e.getScreenY());
                });
    }

    private Menu buildCategoryMenu(LocalDate date, String member, CellState st) {
        Menu category = new Menu("カテゴリを選択");
        String current = st != null ? st.dayPreset : PRESET_WORK;
        for (PresetMenuOption opt : PRESET_MENU_OPTIONS) {
            MenuItem item = new MenuItem(opt.label());
            if (opt.preset().equals(current)) {
                item.setDisable(true);
            } else {
                item.setOnAction(ev -> applyPreset(date, member, opt.preset()));
            }
            category.getItems().add(item);
        }
        return category;
    }

    private void showCellContextMenu(
            Button anchor, LocalDate date, String member, double screenX, double screenY) {
        String dKey = date.toString();
        Map<String, CellState> row = cells.get(dKey);
        CellState st = row != null ? row.get(member) : null;
        boolean commentPresent = st != null && hasComment(st.comment);
        boolean hasHourly = st != null && !st.hourly.isEmpty();
        ContextMenu menu = new ContextMenu();
        menu.getItems().add(buildCategoryMenu(date, member, st));
        MenuItem edit = new MenuItem("コメントを入力…");
        edit.setOnAction(ev -> openCommentDialog(date, member));
        MenuItem delete = new MenuItem("コメントを削除");
        delete.setDisable(!commentPresent);
        delete.setOnAction(ev -> applyComment(date, member, ""));
        MenuItem clearHourly = new MenuItem("時間別編集の内容をクリア");
        clearHourly.setDisable(!hasHourly);
        clearHourly.setOnAction(ev -> clearHourlyEdit(date, member));
        menu.getItems()
                .addAll(
                        new SeparatorMenuItem(),
                        edit,
                        delete,
                        new SeparatorMenuItem(),
                        clearHourly);
        menu.setOnHidden(ev -> releaseCellFocus(anchor));
        menu.show(anchor, screenX, screenY);
    }

    private void releaseCellFocus(Button cell) {
        if (cell != null && cell.isFocused()) {
            scroll.requestFocus();
        }
    }

    private void openCommentDialog(LocalDate date, String member) {
        String dKey = date.toString();
        Map<String, CellState> row = cells.get(dKey);
        CellState st = row != null ? row.get(member) : null;
        String initial = st != null ? st.comment : "";
        Stage owner =
                commentDialogOwner instanceof Stage s
                        ? s
                        : commentDialogOwner != null && commentDialogOwner.getScene() != null
                                && commentDialogOwner.getScene().getWindow() instanceof Stage s2
                                ? s2
                                : null;
        MemberAttendanceCellCommentDialog.show(
                owner,
                member,
                dKey,
                initial,
                text -> applyComment(date, member, text));
    }

    private void applyComment(LocalDate date, String member, String comment) {
        String dKey = date.toString();
        String norm = comment != null ? comment.strip() : "";
        CellState prev =
                cells.computeIfAbsent(dKey, k -> new HashMap<>())
                        .getOrDefault(
                                member,
                                new CellState(
                                        PRESET_WORK,
                                        "通常",
                                        companyKindFor(date, dKey),
                                        false,
                                        Map.of(),
                                        ""));
        if (norm.equals(prev.comment.strip())) {
            CellUi ui = cellUiMap.get(cellKey(dKey, member));
            if (ui != null) {
                releaseCellFocus(ui.button());
            }
            return;
        }
        commentEditedCellKeys.add(cellKey(dKey, member));
        cells.get(dKey)
                .put(
                        member,
                        new CellState(
                                prev.dayPreset,
                                prev.leaveType,
                                prev.companyKind,
                                true,
                                new HashMap<>(prev.hourly),
                                norm));
        updateCellButton(dKey, member);
        notifyDirtyChanged();
        CellUi ui = cellUiMap.get(cellKey(dKey, member));
        if (ui != null) {
            releaseCellFocus(ui.button());
        }
    }

    private static void applyCellTooltip(Button cell, CellState st) {
        if (st == null) {
            cell.setTooltip(null);
            return;
        }
        String preset = presetLabel(st.dayPreset, st.leaveType);
        StringBuilder tip = new StringBuilder(preset);
        if (st.companyKind != null && !st.companyKind.isBlank()) {
            tip.append("（会社: ").append(companyKindLabel(st.companyKind)).append("）");
        }
        if (hasComment(st.comment)) {
            tip.append("\nコメント: ").append(st.comment.trim());
        }
        if (!st.hourly.isEmpty()) {
            tip.append("\n時間別編集あり");
        }
        cell.setTooltip(new Tooltip(tip.toString()));
    }

    private static String presetLabel(String dayPreset, String leaveType) {
        if (leaveType != null && !leaveType.isBlank()) {
            return leaveType;
        }
        return switch (dayPreset != null ? dayPreset : PRESET_WORK) {
            case PRESET_OFF_FULL -> "全休";
            case PRESET_PAID_LEAVE -> "有給休暇(年休)";
            case PRESET_ABSENT -> "欠勤";
            case PRESET_OFF_AM -> "前休";
            case PRESET_OFF_PM -> "後休";
            case PRESET_HOLIDAY_WORK -> "休出（休日出勤）";
            case PRESET_HOLIDAY_WORK_AM -> "午前休出";
            case PRESET_HOLIDAY_WORK_PM -> "午後休出";
            case PRESET_NO_DISPATCH -> "配台外";
            default -> "通常";
        };
    }

    private static String companyKindLabel(String kind) {
        if (KIND_PUBLIC.equals(kind)) {
            return "公休日";
        }
        if (KIND_SPECIAL.equals(kind)) {
            return "特別休暇";
        }
        return "平日";
    }

    private static Label legendChip(String text, String styleClass) {
        Label chip = new Label(text);
        chip.getStyleClass().add("pm-attendance-legend-chip");
        chip.getStyleClass().add(styleClass);
        return chip;
    }
}
