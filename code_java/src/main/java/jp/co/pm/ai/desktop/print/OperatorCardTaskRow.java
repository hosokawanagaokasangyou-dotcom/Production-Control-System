package jp.co.pm.ai.desktop.print;

public record OperatorCardTaskRow(
        String timeRange,
        String processName,
        String machineName,
        String requestNo,
        String qtyDispatchDay,
        String qtyConverted,
        String memberNames) {}
