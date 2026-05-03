package jp.co.pm.ai.desktop.print;

import java.time.LocalDate;
import java.util.List;

public record OperatorCardDaySection(
        LocalDate date, String dateColumnHeader, List<OperatorCardTaskRow> rows) {}
