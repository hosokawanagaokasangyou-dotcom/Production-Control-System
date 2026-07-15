package jp.co.pm.ai.desktop.ui;

import java.util.List;

/** 「担当OP_限定」編集に必要な行の工程名・機械名。 */
public record LimitedOperatorEditContext(String processName, String machineName) {

    public LimitedOperatorEditContext {
        processName = processName != null ? processName.strip() : "";
        machineName = machineName != null ? machineName.strip() : "";
    }

    public static LimitedOperatorEditContext fromRow(
            List<String> headers, List<String> row) {
        return new LimitedOperatorEditContext(
                cellAt(headers, row, "工程名"),
                cellAt(headers, row, "機械名"));
    }

    public void validateComplete() {
        if (processName.isEmpty() || machineName.isEmpty()) {
            throw new IllegalArgumentException(
                    "対象行の「工程名」と「機械名」を入力してから担当者を選択してください。");
        }
    }

    private static String cellAt(List<String> headers, List<String> row, String columnName) {
        int column = headers != null ? headers.indexOf(columnName) : -1;
        if (column < 0 || row == null || column >= row.size()) {
            return "";
        }
        String value = row.get(column);
        return value != null ? value : "";
    }
}
