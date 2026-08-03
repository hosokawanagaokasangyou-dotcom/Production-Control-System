package jp.co.pm.ai.desktop.ui;

import java.util.List;

/**
 * 「AI特別指定_解析」列。段階2 が「特別指定_備考」の解析結果を行ごとに書き戻す表示専用列で、
 * 手入力しても配台には効かない。
 */
public final class PlanInputAiSpecialParseColumn {

    public static final String COLUMN_TITLE = "AI特別指定_解析";
    public static final String SOURCE_COLUMN_TITLE = "特別指定_備考";

    private PlanInputAiSpecialParseColumn() {}

    /**
     * 備考を編集した行の解析セルを空にする。次の段階2 実行まで、編集前の備考に対する解析を出したままにしない。
     *
     * @return クリアしたとき {@code true}
     */
    public static boolean clearStaleParseAfterRemarkEdit(
            List<String> headers, List<String> row, String editedColumnTitle) {
        if (headers == null || row == null || !SOURCE_COLUMN_TITLE.equals(editedColumnTitle)) {
            return false;
        }
        int parseIndex = headers.indexOf(COLUMN_TITLE);
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
