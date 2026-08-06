package jp.co.pm.ai.desktop.ui;

import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.format.TextStyle;
import java.time.temporal.WeekFields;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

import javafx.application.Platform;
import javafx.geometry.HPos;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.Tooltip;
import javafx.scene.layout.ColumnConstraints;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.Priority;
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

    private FiscalYearPeriod fiscalPeriod = FiscalYearPeriod.DEFAULT_APRIL_MARCH;
    private int fiscalYearLabel =
            FiscalYearPeriod.fiscalYearLabelFor(
                    LocalDate.now(), FiscalYearPeriod.DEFAULT_APRIL_MARCH);

    public EditableCompanyCalendarPane() {
        getStyleClass().add("pm-company-calendar");
        setSpacing(6);

        Label legend =
                new Label("クリックで切替: 出勤日 → 公休 → 特別休暇（祝日はピンク・祝表示）");
        legend.getStyleClass().add("pm-company-calendar-legend");
        legend.setWrapText(true);

        fiscalTitleLabel.getStyleClass().add("pm-company-calendar-fiscal-title");

        yearGrid.setHgap(10);
        yearGrid.setVgap(8);
        yearGrid.setPadding(new Insets(4, 0, 8, 0));

        scroll.setContent(yearGrid);
        scroll.setFitToWidth(true);
        scroll.setStyle("-fx-background-color: transparent;");
        VBox.setVgrow(scroll, Priority.ALWAYS);

        getChildren().addAll(legend, fiscalTitleLabel, scroll);
        rebuildYearGrid();
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
}
