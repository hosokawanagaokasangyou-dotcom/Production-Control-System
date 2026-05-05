package jp.co.pm.ai.desktop.ui;

import javafx.geometry.Point2D;
import javafx.scene.SnapshotParameters;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableRow;
import javafx.scene.effect.BlurType;
import javafx.scene.effect.DropShadow;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.input.Dragboard;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.Pane;
import javafx.scene.paint.Color;

/**
 * Sets {@link Dragboard#setDragView} to a snapshot of the full {@link TableRow} with shadow (Spreadsheet row reorder).
 */
public final class SpreadsheetRowReorderDragGhost {

    private static final double PAD = 20.0;

    private SpreadsheetRowReorderDragGhost() {}

    public static void apply(Dragboard db, TableCell<?, ?> tc, MouseEvent e) {
        TableRow<?> row = tc.getTableRow();
        if (row == null) {
            return;
        }
        double rw = row.getWidth();
        double rh = row.getHeight();
        if (rw <= 1 || rh <= 1) {
            return;
        }
        try {
            SnapshotParameters rowParams = new SnapshotParameters();
            rowParams.setFill(Color.TRANSPARENT);
            Image base = row.snapshot(rowParams, null);
            if (base == null) {
                return;
            }

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
            Image ghost = plate.snapshot(outParams, null);
            if (ghost == null) {
                Point2D fb = row.sceneToLocal(e.getSceneX(), e.getSceneY());
                db.setDragView(base, clamp(fb.getX(), 0, base.getWidth()), clamp(fb.getY(), 0, base.getHeight()));
                return;
            }

            Point2D local = row.sceneToLocal(e.getSceneX(), e.getSceneY());
            double ox = clamp(local.getX(), 0, base.getWidth()) + PAD;
            double oy = clamp(local.getY(), 0, base.getHeight()) + PAD;
            ox = clamp(ox, 0, ghost.getWidth());
            oy = clamp(oy, 0, ghost.getHeight());
            db.setDragView(ghost, ox, oy);
        } catch (RuntimeException ignored) {
            // default drag appearance
        }
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
