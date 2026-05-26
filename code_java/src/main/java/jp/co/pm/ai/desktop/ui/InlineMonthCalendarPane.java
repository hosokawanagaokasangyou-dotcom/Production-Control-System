package jp.co.pm.ai.desktop.ui;

import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.time.format.TextStyle;
import java.time.temporal.WeekFields;
import java.util.Locale;
import java.util.Objects;

import javafx.beans.property.ObjectProperty;
import javafx.beans.property.SimpleObjectProperty;
import javafx.geometry.HPos;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.layout.ColumnConstraints;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;

/** 月単位の常時表示カレンダー（単一選択）。 */
public final class InlineMonthCalendarPane extends VBox {

    private static final DateTimeFormatter MONTH_TITLE =
            DateTimeFormatter.ofPattern("yyyy年 M月", Locale.JAPANESE);

    private final ObjectProperty<LocalDate> selectedDate =
            new SimpleObjectProperty<>(LocalDate.now());
    private final ObjectProperty<YearMonth> displayedMonth =
            new SimpleObjectProperty<>(YearMonth.now());

    private final Label monthLabel = new Label();
    private final GridPane dayGrid = new GridPane();

    public InlineMonthCalendarPane() {
        getStyleClass().add("pm-inline-month-calendar");
        setSpacing(6);

        Button prev = new Button("◀");
        Button next = new Button("▶");
        prev.getStyleClass().add("pm-inline-month-calendar-nav");
        next.getStyleClass().add("pm-inline-month-calendar-nav");
        monthLabel.getStyleClass().add("pm-inline-month-calendar-title");
        Region spacer = new Region();
        HBox.setHgrow(spacer, Priority.ALWAYS);
        HBox header = new HBox(8, prev, monthLabel, spacer, next);
        header.setAlignment(Pos.CENTER_LEFT);
        prev.setOnAction(e -> displayedMonth.set(displayedMonth.get().minusMonths(1)));
        next.setOnAction(e -> displayedMonth.set(displayedMonth.get().plusMonths(1)));

        dayGrid.setHgap(4);
        dayGrid.setVgap(4);
        dayGrid.getStyleClass().add("pm-inline-month-calendar-grid");
        for (int i = 0; i < 7; i++) {
            ColumnConstraints cc = new ColumnConstraints();
            cc.setHgrow(Priority.ALWAYS);
            cc.setMinWidth(28);
            dayGrid.getColumnConstraints().add(cc);
        }

        getChildren().addAll(header, dayGrid);

        displayedMonth.addListener((obs, oldMonth, newMonth) -> rebuildGrid());
        selectedDate.addListener(
                (obs, oldDate, newDate) -> {
                    if (newDate == null) {
                        rebuildGrid();
                        return;
                    }
                    YearMonth month = YearMonth.from(newDate);
                    if (!Objects.equals(month, displayedMonth.get())) {
                        displayedMonth.set(month);
                    } else {
                        rebuildGrid();
                    }
                });

        rebuildGrid();
    }

    public ObjectProperty<LocalDate> selectedDateProperty() {
        return selectedDate;
    }

    public LocalDate getSelectedDate() {
        return selectedDate.get();
    }

    public void setSelectedDate(LocalDate date) {
        selectedDate.set(date);
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
            GridPane.setHalignment(weekday, HPos.CENTER);
            dayGrid.add(weekday, col, 0);
        }

        LocalDate first = month.atDay(1);
        int offset =
                (first.getDayOfWeek().getValue() - firstDayOfWeek.getValue() + 7) % 7;
        int daysInMonth = month.lengthOfMonth();
        LocalDate today = LocalDate.now();
        LocalDate selected = selectedDate.get();

        for (int day = 1; day <= daysInMonth; day++) {
            LocalDate date = month.atDay(day);
            Button cell = new Button(Integer.toString(day));
            cell.getStyleClass().add("pm-inline-month-calendar-day");
            cell.setMaxWidth(Double.MAX_VALUE);
            if (date.equals(today)) {
                cell.getStyleClass().add("pm-inline-month-calendar-day-today");
            }
            if (selected != null && date.equals(selected)) {
                cell.getStyleClass().add("pm-inline-month-calendar-day-selected");
            }
            cell.setOnAction(e -> selectedDate.set(date));
            int index = offset + day - 1;
            dayGrid.add(cell, index % 7, 1 + index / 7);
        }
    }
}
