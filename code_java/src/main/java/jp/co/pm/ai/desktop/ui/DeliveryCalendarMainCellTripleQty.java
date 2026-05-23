package jp.co.pm.ai.desktop.ui;

/**
 * Three stacked quantity lines (plan / actual / dispatch / stage3). UI prefixes in {@link
 * SpreadsheetTabularSupport#buildReadOnlyDeliveryCalendarMainGrid}. Top-level record: nested record .class can be
 * missing under javafx:run (see {@link Stage2PlanRowDispatchQtyMetricsResult}).
 */
public record DeliveryCalendarMainCellTripleQty(
        String plan, String actual, String dispatch, String stage3After)
        implements DeliveryCalendarMainCell {}
