package jp.co.pm.ai.desktop.ui;

/**
 * One main-grid cell: plain string columns / section headers, or three stacked numeric lines (task-input
 * Aladdin, actual detail aggregate, result dispatch table JSON). UI prefixes lines as {@code (????)},
 * {@code (??)}, {@code (????)} in {@link SpreadsheetTabularSupport#buildReadOnlyDeliveryCalendarMainGrid}.
 */
public sealed interface DeliveryCalendarMainCell
        permits DeliveryCalendarMainCellPlainText, DeliveryCalendarMainCellTripleQty {}
