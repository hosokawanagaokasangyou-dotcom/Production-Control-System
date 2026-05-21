package jp.co.pm.ai.desktop.print;

import java.util.List;

/** One printable card for a single operator (consecutive calendar days). */
public record OperatorCardPage(String operatorName, List<OperatorCardDaySection> days) {}
