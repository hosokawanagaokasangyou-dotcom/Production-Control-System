package jp.co.pm.ai.desktop.ui;

import java.time.DayOfWeek;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.Map;

import com.fasterxml.jackson.databind.JsonNode;

import javafx.scene.control.Button;
import javafx.scene.control.Tooltip;

/** 会社カレンダー日セルの表示（InlineMonthCalendarPane / EditableCompanyCalendarPane 共通）。 */
public final class CompanyCalendarDayVisual {

    public static final String KIND_WORKING = "working_day";
    public static final String KIND_PUBLIC = "public_holiday";
    public static final String KIND_SPECIAL = "special_holiday";
    public static final String SOURCE_NATIONAL = "national_holiday";

    public record DayInfo(String kind, String label, boolean manualEdit, String source) {}

    private CompanyCalendarDayVisual() {}

    public static Map<LocalDate, DayInfo> parseDays(JsonNode daysNode) {
        Map<LocalDate, DayInfo> out = new HashMap<>();
        if (daysNode == null || !daysNode.isObject()) {
            return out;
        }
        daysNode
                .fields()
                .forEachRemaining(
                        e -> {
                            try {
                                LocalDate d = LocalDate.parse(e.getKey());
                                JsonNode v = e.getValue();
                                out.put(
                                        d,
                                        new DayInfo(
                                                v.path("kind").asText(""),
                                                v.path("label").asText(""),
                                                v.path("manual_edit").asBoolean(false),
                                                v.has("source") ? v.path("source").asText("") : null));
                            } catch (Exception ignored) {
                                // skip invalid keys
                            }
                        });
        return out;
    }

    public static String defaultKindFor(LocalDate date) {
        if (date.getDayOfWeek() == DayOfWeek.SATURDAY || date.getDayOfWeek() == DayOfWeek.SUNDAY) {
            return KIND_PUBLIC;
        }
        return KIND_WORKING;
    }

    public static String shortLabel(int day, String kind, DayInfo entry) {
        if (KIND_PUBLIC.equals(kind)) {
            return day + "公";
        }
        if (KIND_SPECIAL.equals(kind)) {
            return day + "特";
        }
        return Integer.toString(day);
    }

    public static String kindLabel(String kind) {
        return switch (kind) {
            case KIND_PUBLIC -> "公休";
            case KIND_SPECIAL -> "特別休暇";
            default -> "出勤";
        };
    }

    public static void applyToDayButton(Button cell, LocalDate date, DayInfo entry) {
        String kind = entry != null ? entry.kind : defaultKindFor(date);
        cell.setText(shortLabel(date.getDayOfMonth(), kind, entry));
        cell.getStyleClass().add("pm-company-calendar-day");
        applyCellStyle(cell, kind, entry);
        if (entry != null && entry.label != null && !entry.label.isBlank()) {
            cell.setTooltip(new Tooltip(entry.label));
        } else {
            cell.setTooltip(new Tooltip(kindLabel(kind)));
        }
        cell.setAccessibleText(
                date
                        + " · "
                        + kindLabel(kind)
                        + (entry != null && entry.label != null ? " " + entry.label : ""));
    }

    private static void applyCellStyle(Button cell, String kind, DayInfo entry) {
        cell.getStyleClass()
                .removeAll(
                        "pm-company-cal-public",
                        "pm-company-cal-special",
                        "pm-company-cal-working",
                        "pm-company-cal-national");
        switch (kind) {
            case KIND_PUBLIC -> cell.getStyleClass().add("pm-company-cal-public");
            case KIND_SPECIAL -> cell.getStyleClass().add("pm-company-cal-special");
            default -> cell.getStyleClass().add("pm-company-cal-working");
        }
    }
}
