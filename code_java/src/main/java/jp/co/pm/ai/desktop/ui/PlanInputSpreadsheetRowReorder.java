package jp.co.pm.ai.desktop.ui;

import java.util.List;

import javafx.collections.ObservableList;

/**
 * 配台計画_タスク入力／配台計画_タスク入力3.0 共通の行並べ替え後処理。
 *
 * <ul>
 *   <li>入力3.0: 元依頼NO単位で全枝番を相対順のまま移動（枝番順・試行順連続を維持）
 *   <li>入力1表: 同一依頼NOの配台対象行をブロックで移動（配台不要=yes は追従しない）
 *   <li>↑↓ は隣接行ではなく隣接<strong>ブロック</strong>と入れ替え（DnD と同じ単位）
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
     * ↑ ボタン。選択データ行 index を 1 段上へ（複数工程・枝番はブロック単位）。
     *
     * @return 移動後のフォーカス行 index
     */
    public static int moveBlockUp(
            List<String> headers,
            ObservableList<ObservableList<String>> rows,
            int selectedDataIndex) {
        int focus =
                PlanInputProcessSequenceRowOrder.moveBlockOneStepUp(
                        headers, rows, selectedDataIndex);
        stabilizeAndRenumberDispatchTrialOrder(headers, rows);
        return focus;
    }

    /**
     * ↓ ボタン。選択データ行 index を 1 段下へ（複数工程・枝番はブロック単位）。
     *
     * @return 移動後のフォーカス行 index
     */
    public static int moveBlockDown(
            List<String> headers,
            ObservableList<ObservableList<String>> rows,
            int selectedDataIndex) {
        int focus =
                PlanInputProcessSequenceRowOrder.moveBlockOneStepDown(
                        headers, rows, selectedDataIndex);
        stabilizeAndRenumberDispatchTrialOrder(headers, rows);
        return focus;
    }
}
