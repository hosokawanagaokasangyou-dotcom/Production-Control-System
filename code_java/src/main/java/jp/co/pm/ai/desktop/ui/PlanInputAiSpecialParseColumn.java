package jp.co.pm.ai.desktop.ui;

import java.util.List;

/**
 * 「AI特別指定_解析」列。段階2 が「特別指定_備考」の解析結果を行ごとに書き戻す表示専用列で、
 * 手入力しても配台には効かない。
 */
public final class PlanInputAiSpecialParseColumn {

    public static final String COLUMN_TITLE = "AI特別指定_解析";
    public static final String SOURCE_COLUMN_TITLE = "特別指定_備考";

    /** 工場別に見出しを変えたブック向け（Python {@code PLAN_COL_AI_PARSE_ALIASES} と揃える）。 */
    public static final List<String> PARSE_COLUMN_TITLES =
            List.of(COLUMN_TITLE, "AI納期回答_解析");

    public static final List<String> SOURCE_COLUMN_TITLES =
            List.of(SOURCE_COLUMN_TITLE, "納期回答_備考");

    private PlanInputAiSpecialParseColumn() {}

    public static boolean isParseColumn(String columnTitle) {
        if (columnTitle == null || columnTitle.isBlank()) {
            return false;
        }
        return PARSE_COLUMN_TITLES.contains(columnTitle.trim());
    }

    public static boolean isSourceColumn(String columnTitle) {
        if (columnTitle == null || columnTitle.isBlank()) {
            return false;
        }
        return SOURCE_COLUMN_TITLES.contains(columnTitle.trim());
    }

    public static int indexOfParseColumn(List<String> headers) {
        if (headers == null || headers.isEmpty()) {
            return -1;
        }
        for (String title : PARSE_COLUMN_TITLES) {
            int idx = headers.indexOf(title);
            if (idx >= 0) {
                return idx;
            }
        }
        return -1;
    }

    public static String resolveParseColumnTitle(List<String> headers) {
        int idx = indexOfParseColumn(headers);
        return idx >= 0 ? headers.get(idx) : COLUMN_TITLE;
    }

    public static String resolveSourceColumnTitle(List<String> headers) {
        if (headers != null) {
            for (String title : SOURCE_COLUMN_TITLES) {
                if (headers.contains(title)) {
                    return title;
                }
            }
        }
        return SOURCE_COLUMN_TITLE;
    }

    /**
     * 備考を編集した行の解析セルを空にする。次の段階2 実行まで、編集前の備考に対する解析を出したままにしない。
     *
     * @return クリアしたとき {@code true}
     */
    public static boolean clearStaleParseAfterRemarkEdit(
            List<String> headers, List<String> row, String editedColumnTitle) {
        if (headers == null || row == null || !isSourceColumn(editedColumnTitle)) {
            return false;
        }
        int parseIndex = indexOfParseColumn(headers);
        if (parseIndex < 0 || parseIndex >= row.size()) {
            return false;
        }
        String current = row.get(parseIndex);
        if (current == null || current.isBlank()) {
            return false;
        }
        row.set(parseIndex, "");
        return true;
    }
}
