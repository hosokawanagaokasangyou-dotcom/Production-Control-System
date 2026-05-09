package jp.co.pm.ai.desktop.ui;

/**
 * One main-grid cell: plain string columns / section headers, or three stacked numeric lines (task-input
 * Aladdin, actual detail aggregate, result dispatch table JSON).
 */
public sealed interface DeliveryCalendarMainCell
        permits DeliveryCalendarMainCell.PlainText, DeliveryCalendarMainCell.TripleQty {

    record PlainText(String text) implements DeliveryCalendarMainCell {}

    record TripleQty(String plan, String actual, String dispatch) implements DeliveryCalendarMainCell {}
}
