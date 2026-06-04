package jp.co.pm.ai.desktop.ui;

import javafx.geometry.Point2D;
import javafx.scene.SnapshotParameters;
import javafx.scene.control.TableCell;
import javafx.scene.effect.BlurType;
import javafx.scene.effect.DropShadow;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.input.Dragboard;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.Pane;
import javafx.scene.paint.Color;

/**
 * Sets {@link Dragboard#setDragView} for Spreadsheet row reorder.
 *
 * <p>既定はクリックした {@link TableCell} のみをスナップショットする。{@link TableRow} 全体の snapshot は
 * ControlsFX {@link org.controlsfx.control.spreadsheet.SpreadsheetView} のレイアウトを揺らし、ホスト
 * {@code layoutBounds} 経由で列固定 chrome の再適用が走ってウィンドウサイズが変わったように見えることがある。
 */
public final class SpreadsheetRowReorderDragGhost {

    private static final double PAD = 20.0;

    private SpreadsheetRowReorderDragGhost() {}

    /** 行 DnD 用ドラッグイメージ（セル単位・レイアウト影響を最小化）。 */
    public static void apply(Dragboard db, TableCell<?, ?> tc, MouseEvent e) {
        applyCellGhost(db, tc, e);
    }

    private static void applyCellGhost(Dragboard db, TableCell<?, ?> tc, MouseEvent e) {
        if (tc == null || tc.isEmpty()) {
            return;
        }
        double cw = tc.getWidth();
        double ch = tc.getHeight();
        if (cw <= 1 || ch <= 1) {
            return;
        }
        try {
            SnapshotParameters params = new SnapshotParameters();
            params.setFill(Color.TRANSPARENT);
            Image base = tc.snapshot(params, null);
            if (base == null) {
                return;
            }
            Image ghost = withShadowPlate(base);
            if (ghost == null) {
                Point2D local = tc.sceneToLocal(e.getSceneX(), e.getSceneY());
                db.setDragView(
                        base,
                        clamp(local.getX(), 0, base.getWidth()),
                        clamp(local.getY(), 0, base.getHeight()));
                return;
            }
            Point2D local = tc.sceneToLocal(e.getSceneX(), e.getSceneY());
            double ox = clamp(local.getX(), 0, base.getWidth()) + PAD;
            double oy = clamp(local.getY(), 0, base.getHeight()) + PAD;
            ox = clamp(ox, 0, ghost.getWidth());
            oy = clamp(oy, 0, ghost.getHeight());
            db.setDragView(ghost, ox, oy);
        } catch (RuntimeException ignored) {
            // default drag appearance
        }
    }

    private static Image withShadowPlate(Image base) {
        ImageView iv = new ImageView(base);
        iv.setOpacity(1.0);
        iv.setSmooth(true);

        DropShadow shadow = new DropShadow();
        shadow.setBlurType(BlurType.GAUSSIAN);
        shadow.setRadius(12);
        shadow.setSpread(0.12);
        shadow.setOffsetX(2);
        shadow.setOffsetY(5);
        shadow.setColor(Color.color(0, 0, 0, 0.48));
        iv.setEffect(shadow);

        Pane plate = new Pane(iv);
        iv.setLayoutX(PAD);
        iv.setLayoutY(PAD);
        double pw = base.getWidth() + 2 * PAD;
        double ph = base.getHeight() + 2 * PAD;
        plate.setMinSize(pw, ph);
        plate.setPrefSize(pw, ph);
        plate.setMaxSize(pw, ph);

        SnapshotParameters outParams = new SnapshotParameters();
        outParams.setFill(Color.TRANSPARENT);
        return plate.snapshot(outParams, null);
    }

    private static double clamp(double v, double lo, double hi) {
        if (v < lo) {
            return lo;
        }
        if (v > hi) {
            return hi;
        }
        return v;
    }
}
