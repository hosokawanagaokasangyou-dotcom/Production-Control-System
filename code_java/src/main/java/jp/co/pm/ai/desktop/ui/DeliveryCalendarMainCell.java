package jp.co.pm.ai.desktop.ui;

/**
 * One main-grid cell: plain string columns / section headers, or three stacked numeric lines (task-input
 * Aladdin, actual detail aggregate, result dispatch table JSON). UI prefixes lines as {@code (????)},
 * {@code (??)}, {@code (????)} in {@link SpreadsheetTabularSupport#buildReadOnlyDeliveryCalendarMainGrid}.
 */
public sealed interface DeliveryCalendarMainCell
        permits DeliveryCalendarMainCell.PlainText, DeliveryCalendarMainCell.TripleQty {

    record PlainText(String text) implements DeliveryCalendarMainCell {}

    /**
     * Raw quantity strings from delivery-calendar JSON {@code triple.p/a/d}; UI adds prefixes per line.
     * Blank / zero lines may be omitted in the spreadsheet for readability (see {@link
     * SpreadsheetTabularSupport#buildReadOnlyDeliveryCalendarMainGrid}).
     */
    record TripleQty(String plan, String actual, String dispatch) implements DeliveryCalendarMainCell {}
}
