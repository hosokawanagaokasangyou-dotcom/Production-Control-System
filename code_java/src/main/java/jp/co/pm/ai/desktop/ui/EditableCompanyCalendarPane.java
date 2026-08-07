package jp.co.pm.ai.desktop.ui;

import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.format.TextStyle;
import java.time.temporal.WeekFields;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.function.Consumer;

import javafx.application.Platform;
import javafx.geometry.HPos;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Node;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.Tooltip;
import javafx.scene.layout.ColumnConstraints;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;

/** 会社カレンダー編集用の年度グリッド（公休・特別休暇・出勤日の塗りつぶし）。 */
public final class EditableCompanyCalendarPane extends VBox {

    public static final String KIND_WORKING = "working_day";
    public static final String KIND_PUBLIC = "public_holiday";
    public static final String KIND_SPECIAL = "special_holiday";

    public static final String SOURCE_NATIONAL = "national_holiday";

    private static final int MONTHS_PER_ROW = 4;
    private static final Locale CALENDAR_LOCALE = Locale.JAPAN;

    private int cellSizePx = AttendanceGridCellSizing.DEFAULT_CELL_PX;

    private static final String[] KIND_CYCLE =
            new String[] {KIND_WORKING, KIND_PUBLIC, KIND_SPECIAL};

    private final Map<LocalDate, DayEntry> dayEntries = new HashMap<>();
    private final Map<LocalDate, Button> cellButtons = new HashMap<>();
    private final Label fiscalTitleLabel = new Label();
    private final GridPane yearGrid = new GridPane();
    private final ScrollPane scroll = new ScrollPane();
    private final AttendanceGridLoadingOverlay loadingOverlay =
            new AttendanceGridLoadingOverlay("pm-company-calendar-grid-loading-overlay");
    private final StackPane scrollHost = new StackPane();
    private final GridRowHoverDimmingController rowDimming = new GridRowHoverDimmingController();
    private int nextRowDimmingIndex = 0;
    private Consumer<Boolean> dirtyListener;
    /** 読込／保存直後の export スナップショット（JSON 正本との差分で未保存を判定）。 */
    private Map<String, Map<String, Object>> savedBaseline = Map.of();

    private FiscalYearPeriod fiscalPeriod = FiscalYearPeriod.DEFAULT_APRIL_MARCH;
    private int fiscalYearLabel =
            FiscalYearPeriod.fiscalYearLabelFor(
                    LocalDate.now(), FiscalYearPeriod.DEFAULT_APRIL_MARCH);

    public EditableCompanyCalendarPane() {
        getStyleClass().add("pm-company-calendar");
        setSpacing(6);

        Label legendHint =
                new Label(
                        "クリックで切替　祝日・週末の一括取得はツールバー「セットアップ」");
        legendHint.getStyleClass().add("pm-company-calendar-legend");
        legendHint.setWrapText(true);
        HBox legendChips =
                new HBox(
                        6,
                        legendChip("出勤", "pm-att-legend-work"),
                        legendChip("公休", "pm-att-legend-off"),
                        legendChip("祝日", "pm-att-legend-national"),
                        legendChip("特別", "pm-att-legend-partial"));
        legendChips.getStyleClass().add("pm-attendance-legend-chips");

        fiscalTitleLabel.getStyleClass().add("pm-company-calendar-fiscal-title");

        yearGrid.setHgap(10);
        yearGrid.setVgap(8);
        yearGrid.setPadding(new Insets(4, 0, 8, 0));

        scroll.setContent(yearGrid);
        scroll.setFitToWidth(true);
        scroll.setStyle("-fx-background-color: transparent;");

        scrollHost.getChildren().addAll(scroll, loadingOverlay);
        scrollHost.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
        VBox.setVgrow(scrollHost, Priority.ALWAYS);
        rowDimming.installScrollClearOnExit(scroll);

        getChildren().addAll(legendChips, legendHint, fiscalTitleLabel, scrollHost);
        rebuildYearGrid();
    }

    public void refreshRowHoverDimming() {
        rowDimming.refresh();
    }

    /** Python 読込・保存などの処理中にグリッドを暗転する。 */
    public void setGridLoading(boolean loading) {
        setGridLoading(loading, null);
    }

    private String gridLoadingMessage = null;

    public void setGridLoading(boolean loading, String message) {
        gridLoading = loading;
        gridLoadingMessage = message;
        applyGridOverlayState();
    }

    public void setGridLoadingMessage(String message) {
        if (gridLoading) {
            gridLoadingMessage = message;
            loadingOverlay.setMessage(message);
        }
    }

    /**
     * 段階2未準備（会社カレンダー未設定など）のときグリッドを暗転する。
     * 編集は可能（オーバーレイはマウス透過）。
     */
    public void setGridNeedsAttention(boolean needsAttention) {
        gridNeedsAttention = needsAttention;
        applyGridOverlayState();
    }

    private boolean gridLoading = false;
    private boolean gridNeedsAttention = false;

    private void applyGridOverlayState() {
        boolean dim = gridLoading || gridNeedsAttention;
        toggleStyleClass(scrollHost, "pm-company-calendar-grid-loading", dim);
        if (gridLoading) {
            loadingOverlay.setAttentionOnly(false);
            loadingOverlay.setLoading(true, gridLoadingMessage);
            scroll.setDisable(true);
            setAllCellButtonsDisabled(true);
        } else if (gridNeedsAttention) {
            loadingOverlay.setLoading(false);
            loadingOverlay.setAttentionOnly(true);
            scroll.setDisable(false);
            setAllCellButtonsDisabled(false);
        } else {
            loadingOverlay.setLoading(false);
            loadingOverlay.setAttentionOnly(false);
            scroll.setDisable(false);
            setAllCellButtonsDisabled(false);
        }
    }

    private void setAllCellButtonsDisabled(boolean disabled) {
        for (Button cell : cellButtons.values()) {
            cell.setDisable(disabled);
        }
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

    public int cellSizePx() {
        return cellSizePx;
    }

    public void setCellSizePx(int px) {
        int clamped = AttendanceGridCellSizing.clamp(px);
        if (cellSizePx == clamped) {
            return;
        }
        cellSizePx = clamped;
        rebuildYearGrid();
    }

    public void setFiscalYear(int fiscalYearLabel, FiscalYearPeriod period) {
        if (period != null) {
            this.fiscalPeriod = period;
        }
        this.fiscalYearLabel = fiscalYearLabel;
        rebuildYearGrid();
    }

    public void setFiscalYearAndDays(
            int fiscalYearLabel,
            FiscalYearPeriod period,
            Map<String, Map<String, Object>> days) {
        if (period != null) {
            this.fiscalPeriod = period;
        }
        this.fiscalYearLabel = fiscalYearLabel;
        loadDaysFromMap(days);
        rebuildYearGrid();
    }

    public void setDaysFromJson(Map<String, Map<String, Object>> days) {
        loadDaysFromMap(days);
        rebuildYearGrid();
    }

    public void setDirtyListener(Consumer<Boolean> listener) {
        this.dirtyListener = listener;
    }

    public boolean hasUnsavedEdits() {
        Map<String, Map<String, Object>> current =
                exportDaysJsonForFiscalYear(fiscalYearLabel, fiscalPeriod);
        boolean dirty = !exportMapsEqual(current, savedBaseline);
        return dirty;
    }

    public void clearUnsavedEditFlags() {
        captureSavedBaseline();
        notifyDirtyChanged();
    }

    private void captureSavedBaseline() {
        savedBaseline =
                copyExportMap(exportDaysJsonForFiscalYear(fiscalYearLabel, fiscalPeriod));
    }

    private static Map<String, Map<String, Object>> copyExportMap(
            Map<String, Map<String, Object>> src) {
        Map<String, Map<String, Object>> out = new HashMap<>();
        for (var e : src.entrySet()) {
            out.put(e.getKey(), new HashMap<>(e.getValue()));
        }
        return out;
    }

    private static boolean exportMapsEqual(
            Map<String, Map<String, Object>> a, Map<String, Map<String, Object>> b) {
        if (a.size() != b.size()) {
            return false;
        }
        for (var e : a.entrySet()) {
            Map<String, Object> rowB = b.get(e.getKey());
            if (rowB == null || !exportRowEquals(e.getValue(), rowB)) {
                return false;
            }
        }
        return true;
    }

    private static boolean exportRowEquals(Map<String, Object> a, Map<String, Object> b) {
        if (!nz(a.get("kind")).equals(nz(b.get("kind")))) {
            return false;
        }
        if (!nz(a.get("label")).equals(nz(b.get("label")))) {
            return false;
        }
        if (!nz(a.get("source")).equals(nz(b.get("source")))) {
            return false;
        }
        boolean ma = Boolean.TRUE.equals(a.get("manual_edit"));
        boolean mb = Boolean.TRUE.equals(b.get("manual_edit"));
        return ma == mb;
    }

    private static String nz(Object v) {
        return v == null ? "" : String.valueOf(v).trim();
    }

    private void notifyDirtyChanged() {
        if (dirtyListener != null) {
            dirtyListener.accept(hasUnsavedEdits());
        }
    }

    private void loadDaysFromMap(Map<String, Map<String, Object>> days) {
        dayEntries.clear();
        if (days != null) {
            for (var e : days.entrySet()) {
                try {
                    LocalDate d = LocalDate.parse(e.getKey());
                    Map<String, Object> v = e.getValue();
                    String kind =
                            v.get("kind") != null ? String.valueOf(v.get("kind")) : KIND_WORKING;
                    String label = v.get("label") != null ? String.valueOf(v.get("label")) : "";
                    String source =
                            v.get("source") != null ? String.valueOf(v.get("source")) : "";
                    boolean manual =
                            v.get("manual_edit") instanceof Boolean b
                                    ? b
                                    : Boolean.parseBoolean(String.valueOf(v.get("manual_edit")));
                    dayEntries.put(d, new DayEntry(kind, label, manual, source));
                } catch (Exception ignored) {
                    // skip invalid
                }
            }
        }
        captureSavedBaseline();
        notifyDirtyChanged();
    }

    public Map<String, Map<String, Object>> exportDaysJsonForFiscalYear(
            int label, FiscalYearPeriod period) {
        FiscalYearPeriod p = period != null ? period : fiscalPeriod;
        Map<String, Map<String, Object>> out = new HashMap<>();
        LocalDate start = p.startDate(label);
        LocalDate end = p.endDateInclusive(label);
        for (LocalDate d = start; !d.isAfter(end); d = d.plusDays(1)) {
            DayEntry entry = dayEntries.get(d);
            String kind = entry != null ? entry.kind : defaultKindFor(d);
            Map<String, Object> row = new HashMap<>();
            row.put("kind", kind);
            String labelText =
                    entry != null && entry.label != null && !entry.label.isBlank()
                            ? entry.label
                            : defaultLabel(kind);
            if (labelText != null && !labelText.isBlank()) {
                row.put("label", labelText);
            }
            if (entry != null
                    && entry.source != null
                    && !entry.source.isBlank()) {
                row.put("source", entry.source);
            }
            if (entry != null && entry.manualEdit) {
                row.put("manual_edit", true);
            }
            out.put(d.toString(), row);
        }
        return out;
    }

    public FiscalYearPeriod fiscalPeriod() {
        return fiscalPeriod;
    }

    public int fiscalYearLabel() {
        return fiscalYearLabel;
    }

    private static Map<String, Object> exportDayEntry(DayEntry entry) {
        Map<String, Object> row = new HashMap<>();
        row.put("kind", entry.kind);
        if (entry.label != null && !entry.label.isBlank()) {
            row.put("label", entry.label);
        }
        if (entry.source != null && !entry.source.isBlank()) {
            row.put("source", entry.source);
        }
        row.put("manual_edit", true);
        return row;
    }

    private void rebuildYearGrid() {
        double scrollPos = scroll.getVvalue();
        yearGrid.getChildren().clear();
        cellButtons.clear();
        rowDimming.clear();
        nextRowDimmingIndex = 0;
        fiscalTitleLabel.setText(fiscalPeriod.rangeLabel(fiscalYearLabel));

        int col = 0;
        int row = 0;
        for (YearMonth month : fiscalPeriod.monthsInOrder(fiscalYearLabel)) {
            yearGrid.add(buildMonthPanel(month), col, row);
            col++;
            if (col >= MONTHS_PER_ROW) {
                col = 0;
                row++;
            }
        }
        Platform.runLater(() -> scroll.setVvalue(scrollPos));
    }

    private VBox buildMonthPanel(YearMonth month) {
        Label monthTitle =
                new Label(
                        month.getMonth().getDisplayName(TextStyle.SHORT, CALENDAR_LOCALE)
                                + "（"
                                + month.getYear()
                                + "）");
        monthTitle.getStyleClass().add("pm-company-calendar-month-title");

        GridPane dayGrid = new GridPane();
        dayGrid.setHgap(2);
        dayGrid.setVgap(2);
        for (int i = 0; i < 7; i++) {
            ColumnConstraints cc =
                    new ColumnConstraints(
                            cellSizePx, cellSizePx, cellSizePx);
            cc.setHgrow(Priority.NEVER);
            dayGrid.getColumnConstraints().add(cc);
        }

        WeekFields weekFields = WeekFields.of(CALENDAR_LOCALE);
        DayOfWeek firstDayOfWeek = weekFields.getFirstDayOfWeek();
        for (int c = 0; c < 7; c++) {
            DayOfWeek dow = firstDayOfWeek.plus(c);
            Label weekday =
                    new Label(
                            dow.getDisplayName(TextStyle.NARROW_STANDALONE, CALENDAR_LOCALE));
            weekday.getStyleClass().add("pm-company-calendar-weekday");
            AttendanceGridCellSizing.applyHeaderLabel(weekday, cellSizePx);
            GridPane.setHalignment(weekday, HPos.CENTER);
            dayGrid.add(weekday, c, 0);
        }

        LocalDate first = month.atDay(1);
        int offset =
                (first.getDayOfWeek().getValue() - firstDayOfWeek.getValue() + 7) % 7;
        int daysInMonth = month.lengthOfMonth();
        LocalDate today = LocalDate.now();
        LocalDate fiscalStart = fiscalPeriod.startDate(fiscalYearLabel);
        LocalDate fiscalEnd = fiscalPeriod.endDateInclusive(fiscalYearLabel);

        for (int day = 1; day <= daysInMonth; day++) {
            LocalDate date = month.atDay(day);
            if (date.isBefore(fiscalStart) || date.isAfter(fiscalEnd)) {
                continue;
            }
            DayEntry entry = dayEntries.get(date);
            String kind = entry != null ? entry.kind : defaultKindFor(date);
            Button cell = new Button(shortLabel(day, kind, entry));
            cell.getStyleClass().add("pm-inline-month-calendar-day");
            cell.getStyleClass().add("pm-company-calendar-day");
            applyCellStyle(cell, kind, entry);
            bindSquareCell(cell, cellSizePx);
            if (date.equals(today)) {
                cell.getStyleClass().add("pm-inline-month-calendar-day-today");
            }
            applyCellTooltip(cell, entry, kind);
            cell.setAccessibleText(
                    date
                            + " · "
                            + kindLabel(kind)
                            + (entry != null && entry.label != null ? " " + entry.label : ""));
            cell.setOnAction(e -> cycleDay(date));
            cellButtons.put(date, cell);
            int index = offset + day - 1;
            dayGrid.add(cell, index % 7, 1 + index / 7);
        }

        for (int weekRow = 0; weekRow < 6; weekRow++) {
            int gridRow = 1 + weekRow;
            List<Node> weekNodes = new ArrayList<>();
            for (Node child : dayGrid.getChildren()) {
                Integer r = GridPane.getRowIndex(child);
                if (r != null && r == gridRow) {
                    weekNodes.add(child);
                }
            }
            if (weekNodes.isEmpty()) {
                continue;
            }
            Region band = new Region();
            band.getStyleClass().add(GridRowHoverDimmingController.STYLE_BAND);
            band.setMouseTransparent(true);
            band.setMaxWidth(Double.MAX_VALUE);
            band.setMaxHeight(Double.MAX_VALUE);
            GridPane.setColumnSpan(band, 7);
            dayGrid.add(band, 0, gridRow);
            int dimIdx = nextRowDimmingIndex++;
            rowDimming.addRow(band, null, weekNodes);
            for (Node n : weekNodes) {
                rowDimming.installHover(n, dimIdx);
            }
        }

        VBox panel = new VBox(2, monthTitle, dayGrid);
        panel.getStyleClass().add("pm-company-calendar-month-panel");
        return panel;
    }

    private void cycleDay(LocalDate date) {
        DayEntry entry = dayEntries.get(date);
        String current = entry != null ? entry.kind : defaultKindFor(date);
        String next = nextKind(current);
        dayEntries.put(
                date,
                new DayEntry(next, defaultLabel(next), true, null));
        updateCellButton(date);
        notifyDirtyChanged();
    }

    private static String nextKind(String current) {
        for (int i = 0; i < KIND_CYCLE.length; i++) {
            if (KIND_CYCLE[i].equals(current)) {
                return KIND_CYCLE[(i + 1) % KIND_CYCLE.length];
            }
        }
        return KIND_PUBLIC;
    }

    private void updateCellButton(LocalDate date) {
        Button cell = cellButtons.get(date);
        if (cell == null) {
            return;
        }
        DayEntry entry = dayEntries.get(date);
        String kind = entry != null ? entry.kind : defaultKindFor(date);
        cell.setText(shortLabel(date.getDayOfMonth(), kind, entry));
        applyCellStyle(cell, kind, entry);
        applyCellTooltip(cell, entry, kind);
    }

    private static void applyCellTooltip(Button cell, DayEntry entry, String kind) {
        if (entry != null && entry.label != null && !entry.label.isBlank()) {
            cell.setTooltip(new Tooltip(entry.label));
        } else {
            cell.setTooltip(new Tooltip(kindLabel(kind)));
        }
    }

    private static void bindSquareCell(Button cell, int cellPx) {
        AttendanceGridCellSizing.applySquareCell(cell, cellPx);
    }

    private static String defaultKindFor(LocalDate date) {
        if (date.getDayOfWeek() == DayOfWeek.SATURDAY || date.getDayOfWeek() == DayOfWeek.SUNDAY) {
            return KIND_PUBLIC;
        }
        return KIND_WORKING;
    }

    private static String defaultLabel(String kind) {
        return switch (kind) {
            case KIND_PUBLIC -> "公休";
            case KIND_SPECIAL -> "特別休暇";
            default -> "";
        };
    }

    private static String kindLabel(String kind) {
        return switch (kind) {
            case KIND_PUBLIC -> "公休";
            case KIND_SPECIAL -> "特別休暇";
            default -> "出勤";
        };
    }

    private static String shortLabel(int day, String kind, DayEntry entry) {
        if (KIND_PUBLIC.equals(kind)) {
            if (entry != null && SOURCE_NATIONAL.equals(entry.source)) {
                return day + "祝";
            }
            if (entry != null
                    && entry.label != null
                    && !entry.label.isBlank()
                    && !isGenericPublicLabel(entry.label)) {
                return day + "祝";
            }
            return day + "公";
        }
        if (KIND_SPECIAL.equals(kind)) {
            return day + "特";
        }
        return Integer.toString(day);
    }

    private static boolean isGenericPublicLabel(String label) {
        String t = label.strip();
        return t.equals("公休")
                || t.equals("土曜")
                || t.equals("日曜")
                || t.equals("祝日");
    }

    private static void applyCellStyle(Button cell, String kind, DayEntry entry) {
        cell.getStyleClass()
                .removeAll(
                        "pm-company-cal-public",
                        "pm-company-cal-special",
                        "pm-company-cal-working",
                        "pm-company-cal-national");
        switch (kind) {
            case KIND_PUBLIC -> {
                cell.getStyleClass().add("pm-company-cal-public");
                if (entry != null
                        && (SOURCE_NATIONAL.equals(entry.source)
                                || (entry.label != null
                                        && !entry.label.isBlank()
                                        && !isGenericPublicLabel(entry.label)))) {
                    cell.getStyleClass().add("pm-company-cal-national");
                }
            }
            case KIND_SPECIAL -> cell.getStyleClass().add("pm-company-cal-special");
            default -> cell.getStyleClass().add("pm-company-cal-working");
        }
    }

    private record DayEntry(String kind, String label, boolean manualEdit, String source) {}

    private static Label legendChip(String text, String styleClass) {
        Label chip = new Label(text);
        chip.getStyleClass().add("pm-attendance-legend-chip");
        chip.getStyleClass().add(styleClass);
        return chip;
    }
}
