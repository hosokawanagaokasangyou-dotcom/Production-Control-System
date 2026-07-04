package jp.co.pm.ai.desktop.ui;

import javafx.scene.control.TableCell;
import javafx.scene.input.Dragboard;
import javafx.scene.input.MouseEvent;

/**
 * Sets {@link Dragboard#setDragView} for Spreadsheet row reorder.
 *
 * <p>セル／行の {@code snapshot} は ControlsFX {@link org.controlsfx.control.spreadsheet.SpreadsheetView}
 * のレイアウトを揺らし、ホスト {@code layoutBounds} が大幅に縮む。行並べ替えはプラットフォーム既定のドラッグ表示で十分なため、
 * スナップショットは使わない。
 */
public final class SpreadsheetRowReorderDragGhost {

    private SpreadsheetRowReorderDragGhost() {}

    /** 行 DnD 用。レイアウト影響を避けるため {@code setDragView} は呼ばない。 */
    public static void apply(Dragboard db, TableCell<?, ?> tc, MouseEvent e) {
        // platform default drag feedback
    }
}
