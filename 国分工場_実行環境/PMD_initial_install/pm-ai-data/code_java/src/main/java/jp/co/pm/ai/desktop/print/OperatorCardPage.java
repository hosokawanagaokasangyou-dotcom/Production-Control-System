package jp.co.pm.ai.desktop.print;

import java.util.List;

/** One physical A4 page for a single operator (three consecutive calendar days). */
public record OperatorCardPage(String operatorName, List<OperatorCardDaySection> days) {}
