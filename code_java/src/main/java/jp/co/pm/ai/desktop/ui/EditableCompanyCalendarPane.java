package jp.co.pm.ai.desktop.ui;

import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.time.format.TextStyle;
import java.time.temporal.WeekFields;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;

import javafx.beans.property.ObjectProperty;
import javafx.beans.property.SimpleObjectProperty;
import javafx.geometry.HPos;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.Toggle;
import javafx.scene.control.ToggleButton;
import javafx.scene.control.ToggleGroup;
import javafx.scene.layout.ColumnConstraints;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;

/** 会社カレンダー編集用の月グリッド（公休・特別休暇・出勤日の塗りつぶし）。 */
public final class EditableCompanyCalendarPane extends VBox {

    public static final String KIND_WORKING = "working_day";
    public static final String KIND_PUBLIC = "public_holiday";
    public static final String KIND_SPECIAL = "special_holiday";

    private static final DateTimeFormatter MONTH_TITLE =
            DateTimeFormatter.ofPattern("yyyy年 M月", Locale.JAPANESE);

    private final ObjectProperty<YearMonth> displayedMonth =
            new SimpleObjectProperty<>(YearMonth.now());
    private final Map<LocalDate, DayEntry> dayEntries = new HashMap<>();
    private final ToggleGroup paintGroup = new ToggleGroup();

    private final Label monthLabel = new Label();
    private final GridPane dayGrid = new GridPane();
    private String paintKind = KIND_PUBLIC;

    public EditableCompanyCalendarPane() {
        getStyleClass().add("pm-company-calendar");
        setSpacing(6);

        ToggleButton publicBtn = new ToggleButton("公休");
        publicBtn.setToggleGroup(paintGroup);
        publicBtn.setUserData(KIND_PUBLIC);
        ToggleButton specialBtn = new ToggleButton("特別休暇");
        specialBtn.setToggleGroup(paintGroup);
        specialBtn.setUserData(KIND_SPECIAL);
        ToggleButton workBtn = new ToggleButton("出勤日");
        workBtn.setToggleGroup(paintGroup);
        workBtn.setUserData(KIND_WORKING);
        publicBtn.setSelected(true);
        paintGroup.selectedToggleProperty().addListener(
                (obs, oldT, newT) -> {
                    if (newT != null && newT.getUserData() instanceof String k) {
                        paintKind = k;
                    }
                });

        Button prev = new Button("◀");
        Button next = new Button("▶");
        monthLabel.getStyleClass().add("pm-inline-month-calendar-title");
        Region spacer = new Region();
        HBox.setHgrow(spacer, Priority.ALWAYS);
        HBox header = new HBox(8, prev, monthLabel, spacer, next);
        header.setAlignment(Pos.CENTER_LEFT);
        header.getStyleClass().add("pm-inline-month-calendar-nav");
        prev.setOnAction(e -> displayedMonth.set(displayedMonth.get().minusMonths(1)));
        next.setOnAction(e -> displayedMonth.set(displayedMonth.get().plusMonths(1)));

        HBox brushes = new HBox(8, publicBtn, specialBtn, workBtn);
        brushes.setAlignment(Pos.CENTER_LEFT);

        dayGrid.setHgap(4);
        dayGrid.setVgap(4);
        for (int i = 0; i < 7; i++) {
            ColumnConstraints cc = new ColumnConstraints();
            cc.setHgrow(Priority.ALWAYS);
            cc.setMinWidth(44);
            dayGrid.getColumnConstraints().add(cc);
        }

        getChildren().addAll(brushes, header, dayGrid);
        displayedMonth.addListener((obs, o, n) -> rebuildGrid());
        rebuildGrid();
    }

    public void setDaysFromJson(Map<String, Map<String, Object>> days) {
        dayEntries.clear();
        if (days != null) {
            for (var e : days.entrySet()) {
                try {
                    LocalDate d = LocalDate.parse(e.getKey());
                    Map<String, Object> v = e.getValue();
                    String kind =
                            v.get("kind") != null ? String.valueOf(v.get("kind")) : KIND_WORKING;
                    String label = v.get("label") != null ? String.valueOf(v.get("label")) : "";
                    dayEntries.put(d, new DayEntry(kind, label));
                } catch (Exception ignored) {
                    // skip invalid
                }
            }
        }
        rebuildGrid();
    }

    public Map<String, Map<String, Object>> exportDaysJsonForYear(int year) {
        Map<String, Map<String, Object>> out = new HashMap<>();
        for (var e : dayEntries.entrySet()) {
            if (e.getKey().getYear() != year) {
                continue;
            }
            out.put(e.getKey().toString(), exportDayEntry(e.getValue()));
        }
        return out;
    }

    public void setDisplayedYearMonth(YearMonth ym) {
        if (ym != null) {
            displayedMonth.set(ym);
        }
    }

    public YearMonth displayedYearMonth() {
        return displayedMonth.get();
    }

    public Map<String, Map<String, Object>> exportDaysJson() {
        Map<String, Map<String, Object>> out = new HashMap<>();
        for (var e : dayEntries.entrySet()) {
            out.put(e.getKey().toString(), exportDayEntry(e.getValue()));
        }
        return out;
    }

    private static Map<String, Object> exportDayEntry(DayEntry entry) {
        Map<String, Object> row = new HashMap<>();
        row.put("kind", entry.kind);
        if (entry.label != null && !entry.label.isBlank()) {
            row.put("label", entry.label);
        }
        row.put("manual_edit", true);
        return row;
    }

    public int displayedYear() {
        YearMonth ym = displayedMonth.get();
        return ym != null ? ym.getYear() : LocalDate.now().getYear();
    }

    private void rebuildGrid() {
        dayGrid.getChildren().clear();
        YearMonth month = displayedMonth.get();
        if (month == null) {
            monthLabel.setText("");
            return;
        }
        monthLabel.setText(MONTH_TITLE.format(month.atDay(1)));

        WeekFields weekFields = WeekFields.of(Locale.getDefault());
        DayOfWeek firstDayOfWeek = weekFields.getFirstDayOfWeek();
        Locale locale = Locale.getDefault();
        for (int col = 0; col < 7; col++) {
            DayOfWeek dow = firstDayOfWeek.plus(col);
            Label weekday =
                    new Label(dow.getDisplayName(TextStyle.SHORT_STANDALONE, locale));
            weekday.getStyleClass().add("pm-inline-month-calendar-weekday");
            if (dow == DayOfWeek.SATURDAY || dow == DayOfWeek.SUNDAY) {
                weekday.getStyleClass().add("overtime-sim-weekend-column");
            }
            GridPane.setHalignment(weekday, HPos.CENTER);
            dayGrid.add(weekday, col, 0);
        }

        LocalDate first = month.atDay(1);
        int offset =
                (first.getDayOfWeek().getValue() - firstDayOfWeek.getValue() + 7) % 7;
        int daysInMonth = month.lengthOfMonth();
        LocalDate today = LocalDate.now();

        for (int day = 1; day <= daysInMonth; day++) {
            LocalDate date = month.atDay(day);
            DayEntry entry = dayEntries.get(date);
            String kind = entry != null ? entry.kind : defaultKindFor(date);
            Button cell = new Button(shortLabel(day, entry, kind));
            cell.getStyleClass().add("pm-inline-month-calendar-day");
            cell.setMinSize(44, 44);
            cell.setMaxWidth(Double.MAX_VALUE);
            applyCellStyle(cell, kind);
            if (date.equals(today)) {
                cell.getStyleClass().add("pm-inline-month-calendar-day-today");
            }
            cell.setOnAction(
                    e -> {
                        dayEntries.put(date, new DayEntry(paintKind, defaultLabel(paintKind)));
                        rebuildGrid();
                    });
            int index = offset + day - 1;
            dayGrid.add(cell, index % 7, 1 + index / 7);
        }
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

    private static String shortLabel(int day, DayEntry entry, String kind) {
        if (kind.equals(KIND_PUBLIC)) {
            return day + " 公";
        }
        if (kind.equals(KIND_SPECIAL)) {
            return day + " 特";
        }
        return Integer.toString(day);
    }

    private static void applyCellStyle(Button cell, String kind) {
        cell.getStyleClass().removeAll(
                "pm-company-cal-public", "pm-company-cal-special", "pm-company-cal-working");
        switch (kind) {
            case KIND_PUBLIC -> cell.getStyleClass().add("pm-company-cal-public");
            case KIND_SPECIAL -> cell.getStyleClass().add("pm-company-cal-special");
            default -> cell.getStyleClass().add("pm-company-cal-working");
        }
    }

    private record DayEntry(String kind, String label) {}
}
