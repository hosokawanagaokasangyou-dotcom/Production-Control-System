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
     * Raw quantity strings from delivery-calendar JSON {@code triple.p/a/d/s3}; UI adds prefixes per line.
     * {@code stage3After} is {@code 結果_配台表.json} の {@code 実配台数量}（段階3試行後タイムライン m）。
     * Blank / zero lines may be omitted in the spreadsheet for readability (see {@link
     * SpreadsheetTabularSupport#buildReadOnlyDeliveryCalendarMainGrid}).
     */
    record TripleQty(String plan, String actual, String dispatch, String stage3After)
            implements DeliveryCalendarMainCell {}
}
