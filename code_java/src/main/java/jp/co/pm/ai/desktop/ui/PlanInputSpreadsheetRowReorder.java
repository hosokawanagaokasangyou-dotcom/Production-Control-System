package jp.co.pm.ai.desktop.ui;

import java.util.List;

import javafx.collections.ObservableList;

/**
 * 配台計画_タスク入力／配台計画_タスク入力3.0 共通の行並べ替え後処理。
 *
 * <ul>
 *   <li>配台対象行の DnD・↑↓ は同一依頼NOの配台対象行をブロックで移動（配台不要=yes は追従しない）
 *   <li>整列後 {@link PlanInputProcessSequenceRowOrder#COL_DISPATCH_TRIAL_ORDER} を 1..n に振り直す
 *   <li>§A-1（加工内容のカンマ区切り順）を同一依頼NO内で維持
 * </ul>
 */
public final class PlanInputSpreadsheetRowReorder {

    private PlanInputSpreadsheetRowReorder() {}

    /** DnD・↑↓・保存直前など、行順変更後の §A-1 整列と試行順連番。 */
    public static void stabilizeAndRenumberDispatchTrialOrder(
            List<String> headers, ObservableList<ObservableList<String>> rows) {
        PlanInputProcessSequenceRowOrder.stabilizeAndRenumberDispatchTrialOrder(headers, rows);
    }

    /**
     * ↑↓ 用。データ行 index {@code sourceDataIndex} を {@code targetDataIndex} へ移す（配台対象はブロック移動）。
     */
    public static void moveAdjacentDataRows(
            List<String> headers,
            ObservableList<ObservableList<String>> rows,
            int targetDataIndex,
            int sourceDataIndex) {
        if (headers == null
                || rows == null
                || targetDataIndex < 0
                || sourceDataIndex < 0
                || targetDataIndex >= rows.size()
                || sourceDataIndex >= rows.size()
                || targetDataIndex == sourceDataIndex) {
            return;
        }
        PlanInputProcessSequenceRowOrder.moveRowsForUserReorder(
                headers, rows, sourceDataIndex, targetDataIndex);
        stabilizeAndRenumberDispatchTrialOrder(headers, rows);
    }
}
