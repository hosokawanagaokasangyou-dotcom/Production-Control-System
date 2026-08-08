package jp.co.pm.ai.desktop.ui;

import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.time.format.TextStyle;
import java.time.temporal.WeekFields;
import java.util.Locale;
import java.util.Map;
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

    private final boolean monthOnly;
    private boolean companyCalendarMode = false;
    private Map<LocalDate, CompanyCalendarDayVisual.DayInfo> companyCalendarDays = Map.of();
    private final Label monthLabel = new Label();
    private final Button prevButton = new Button("◀");
    private final Button nextButton = new Button("▶");
    private final GridPane dayGrid = new GridPane();

    public InlineMonthCalendarPane() {
        this(false);
    }

    /** @param monthOnly true のとき日グリッドを出さず、◀▶ で月のみ切り替える */
    public InlineMonthCalendarPane(boolean monthOnly) {
        this.monthOnly = monthOnly;
        getStyleClass().add("pm-inline-month-calendar");
        if (monthOnly) {
            getStyleClass().add("pm-inline-month-calendar-month-only");
        }
        setSpacing(6);

        prevButton.getStyleClass().add("pm-inline-month-calendar-nav");
        nextButton.getStyleClass().add("pm-inline-month-calendar-nav");
        monthLabel.getStyleClass().add("pm-inline-month-calendar-title");
        Region spacer = new Region();
        HBox.setHgrow(spacer, Priority.ALWAYS);
        HBox header = new HBox(8, prevButton, monthLabel, spacer, nextButton);
        header.setAlignment(Pos.CENTER_LEFT);
        prevButton.setOnAction(e -> displayedMonth.set(displayedMonth.get().minusMonths(1)));
        nextButton.setOnAction(e -> displayedMonth.set(displayedMonth.get().plusMonths(1)));

        if (!monthOnly) {
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
        } else {
            getChildren().add(header);
        }

        displayedMonth.addListener(
                (obs, oldMonth, newMonth) -> {
                    syncSelectedDateToDisplayedMonth();
                    refreshView();
                });
        selectedDate.addListener(
                (obs, oldDate, newDate) -> {
                    if (newDate == null) {
                        refreshView();
                        return;
                    }
                    YearMonth month = YearMonth.from(newDate);
                    if (!Objects.equals(month, displayedMonth.get())) {
                        displayedMonth.set(month);
                    } else if (!monthOnly) {
                        refreshView();
                    }
                });

        refreshView();
    }

    /** true のとき日セルを会社カレンダー色・表記で表示する（編集不可・選択のみ）。 */
    public void setCompanyCalendarMode(boolean enabled) {
        companyCalendarMode = enabled;
        toggleStyleClass(this, "pm-company-calendar", enabled);
        refreshView();
    }

    public void setCompanyCalendarDays(Map<LocalDate, CompanyCalendarDayVisual.DayInfo> days) {
        companyCalendarDays = days != null ? Map.copyOf(days) : Map.of();
        if (companyCalendarMode && !monthOnly) {
            refreshView();
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

    public ObjectProperty<LocalDate> selectedDateProperty() {
        return selectedDate;
    }

    public LocalDate getSelectedDate() {
        return selectedDate.get();
    }

    public void setSelectedDate(LocalDate date) {
        selectedDate.set(date);
    }

    public ObjectProperty<YearMonth> displayedMonthProperty() {
        return displayedMonth;
    }

    public YearMonth getDisplayedMonth() {
        return displayedMonth.get();
    }

    /** 月ナビ（◀▶）の有効／無効。読込・保存中は false にする。 */
    public void setNavigationEnabled(boolean enabled) {
        prevButton.setDisable(!enabled);
        nextButton.setDisable(!enabled);
    }

    /** ◀▶ で表示月だけ変えたとき、選択日を同月内の同日（月末調整）へ追従する。 */
    private void syncSelectedDateToDisplayedMonth() {
        YearMonth month = displayedMonth.get();
        if (month == null) {
            return;
        }
        LocalDate sel = selectedDate.get();
        if (sel == null) {
            selectedDate.set(month.atDay(1));
            return;
        }
        if (!YearMonth.from(sel).equals(month)) {
            int day =
                    monthOnly
                            ? 1
                            : Math.min(sel.getDayOfMonth(), month.lengthOfMonth());
            selectedDate.set(month.atDay(day));
        }
    }

    private void refreshView() {
        YearMonth month = displayedMonth.get();
        if (month == null) {
            monthLabel.setText("");
            return;
        }
        monthLabel.setText(MONTH_TITLE.format(month.atDay(1)));
        if (!monthOnly) {
            rebuildDayGrid(month);
        }
    }

    private void rebuildDayGrid(YearMonth month) {
        dayGrid.getChildren().clear();

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
            if (companyCalendarMode) {
                CompanyCalendarDayVisual.DayInfo entry = companyCalendarDays.get(date);
                CompanyCalendarDayVisual.applyToDayButton(cell, date, entry);
            }
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
