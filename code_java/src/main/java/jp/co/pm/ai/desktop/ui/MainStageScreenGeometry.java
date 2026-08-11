package jp.co.pm.ai.desktop.ui;

import java.util.List;

import javafx.geometry.Rectangle2D;
import javafx.stage.Screen;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.config.DesktopSessionState;

/**
 * メイン {@link Stage} の位置・サイズをマルチモニター環境で保存・復元する。
 *
 * <p>起動時はプライマリではなく、前回ウィンドウが表示されていたモニターの {@link Screen#getVisualBounds()} 内へ
 * クランプする。
 */
public final class MainStageScreenGeometry {

    private static final double SCREEN_BOUNDS_MATCH_EPSILON = 1.0;

    private MainStageScreenGeometry() {}

  /** セッション保存用のウィンドウ幾何＋表示モニター識別子（visual bounds の min 座標）。 */
    public record Snapshot(
            double width,
            double height,
            double x,
            double y,
            double screenVisualMinX,
            double screenVisualMinY) {

        static Snapshot empty() {
            return new Snapshot(0d, 0d, Double.NaN, Double.NaN, Double.NaN, Double.NaN);
        }
    }

    public static Snapshot fromSessionState(DesktopSessionState state) {
        if (state == null) {
            return Snapshot.empty();
        }
        return new Snapshot(
                state.windowWidth(),
                state.windowHeight(),
                state.windowX(),
                state.windowY(),
                state.windowScreenVisualMinX(),
                state.windowScreenVisualMinY());
    }

    public static Snapshot snapshotFromStage(Stage stage) {
        if (stage == null) {
            return Snapshot.empty();
        }
        double width = stage.getWidth();
        double height = stage.getHeight();
        double x = stage.getX();
        double y = stage.getY();
        Rectangle2D screenBounds = screenBoundsForWindow(x, y, width, height);
        return new Snapshot(
                width,
                height,
                x,
                y,
                screenBounds.getMinX(),
                screenBounds.getMinY());
    }

    public static void applyToStage(Stage stage, Snapshot geometry) {
        if (stage == null || geometry == null) {
            return;
        }
        double minW = stage.getMinWidth();
        double minH = stage.getMinHeight();
        double width = geometry.width();
        double height = geometry.height();
        if (Double.isFinite(width) && Double.isFinite(height) && width >= minW && height >= minH) {
            stage.setWidth(width);
            stage.setHeight(height);
        } else {
            width = stage.getWidth();
            height = stage.getHeight();
        }
        double x = geometry.x();
        double y = geometry.y();
        if (!Double.isFinite(x) || !Double.isFinite(y)) {
            return;
        }
        Rectangle2D screenBounds =
                resolveScreenBounds(
                        geometry.screenVisualMinX(),
                        geometry.screenVisualMinY(),
                        x,
                        y,
                        width,
                        height,
                        collectVisualBounds());
        double maxX = Math.max(screenBounds.getMinX(), screenBounds.getMaxX() - width);
        double maxY = Math.max(screenBounds.getMinY(), screenBounds.getMaxY() - height);
        if (!containsPoint(screenBounds, x, y)) {
            x = screenBounds.getMinX() + (screenBounds.getWidth() - width) / 2.0;
            y = screenBounds.getMinY() + (screenBounds.getHeight() - height) / 2.0;
        }
        stage.setX(clamp(x, screenBounds.getMinX(), maxX));
        stage.setY(clamp(y, screenBounds.getMinY(), maxY));
    }

    /**
     * 単体テスト用。JavaFX {@link Screen} を介さず visual bounds 一覧から対象モニターを解決する。
     */
    static Rectangle2D resolveScreenBoundsForTest(
            Snapshot geometry, List<Rectangle2D> screenVisualBounds) {
        return resolveScreenBounds(
                geometry.screenVisualMinX(),
                geometry.screenVisualMinY(),
                geometry.x(),
                geometry.y(),
                geometry.width(),
                geometry.height(),
                screenVisualBounds);
    }

    private static List<Rectangle2D> collectVisualBounds() {
        return Screen.getScreens().stream().map(Screen::getVisualBounds).toList();
    }

    private static Rectangle2D screenBoundsForWindow(double x, double y, double width, double height) {
        List<Screen> screens = Screen.getScreensForRectangle(x, y, width, height);
        Screen best = screenWithLargestIntersection(screens, x, y, width, height);
        if (best != null) {
            return best.getVisualBounds();
        }
        double cx = x + width / 2.0;
        double cy = y + height / 2.0;
        screens = Screen.getScreensForRectangle(cx, cy, 1, 1);
        if (!screens.isEmpty()) {
            return screens.get(0).getVisualBounds();
        }
        return Screen.getPrimary().getVisualBounds();
    }

    private static Rectangle2D resolveScreenBounds(
            double savedScreenMinX,
            double savedScreenMinY,
            double windowX,
            double windowY,
            double windowWidth,
            double windowHeight,
            List<Rectangle2D> screenVisualBounds) {
        if (screenVisualBounds == null || screenVisualBounds.isEmpty()) {
            return Screen.getPrimary().getVisualBounds();
        }
        if (Double.isFinite(savedScreenMinX) && Double.isFinite(savedScreenMinY)) {
            for (Rectangle2D bounds : screenVisualBounds) {
                if (approxEqual(bounds.getMinX(), savedScreenMinX)
                        && approxEqual(bounds.getMinY(), savedScreenMinY)) {
                    return bounds;
                }
            }
        }
        Rectangle2D bestBounds = boundsWithLargestIntersection(
                screenVisualBounds, windowX, windowY, windowWidth, windowHeight);
        if (bestBounds != null && intersectionArea(bestBounds, windowX, windowY, windowWidth, windowHeight) > 0) {
            return bestBounds;
        }
        double cx = windowX + windowWidth / 2.0;
        double cy = windowY + windowHeight / 2.0;
        for (Rectangle2D bounds : screenVisualBounds) {
            if (containsPoint(bounds, cx, cy)) {
                return bounds;
            }
        }
        for (Rectangle2D bounds : screenVisualBounds) {
            if (bounds.getMinX() == 0 && bounds.getMinY() == 0) {
                return bounds;
            }
        }
        return screenVisualBounds.get(0);
    }

    private static Rectangle2D boundsWithLargestIntersection(
            List<Rectangle2D> screenVisualBounds, double x, double y, double width, double height) {
        if (screenVisualBounds == null || screenVisualBounds.isEmpty()) {
            return null;
        }
        Rectangle2D best = screenVisualBounds.get(0);
        double bestArea = intersectionArea(best, x, y, width, height);
        for (int i = 1; i < screenVisualBounds.size(); i++) {
            Rectangle2D candidate = screenVisualBounds.get(i);
            double area = intersectionArea(candidate, x, y, width, height);
            if (area > bestArea) {
                best = candidate;
                bestArea = area;
            }
        }
        return best;
    }

    private static Screen screenWithLargestIntersection(
            List<Screen> screens, double x, double y, double width, double height) {
        if (screens == null || screens.isEmpty()) {
            return null;
        }
        Screen best = screens.get(0);
        double bestArea = intersectionArea(best.getVisualBounds(), x, y, width, height);
        for (int i = 1; i < screens.size(); i++) {
            Screen candidate = screens.get(i);
            double area = intersectionArea(candidate.getVisualBounds(), x, y, width, height);
            if (area > bestArea) {
                best = candidate;
                bestArea = area;
            }
        }
        return bestArea > 0 ? best : screens.get(0);
    }

    private static double intersectionArea(Rectangle2D screen, double x, double y, double width, double height) {
        double left = Math.max(screen.getMinX(), x);
        double top = Math.max(screen.getMinY(), y);
        double right = Math.min(screen.getMaxX(), x + width);
        double bottom = Math.min(screen.getMaxY(), y + height);
        if (right <= left || bottom <= top) {
            return 0d;
        }
        return (right - left) * (bottom - top);
    }

    private static boolean containsPoint(Rectangle2D bounds, double x, double y) {
        return x >= bounds.getMinX()
                && x < bounds.getMaxX()
                && y >= bounds.getMinY()
                && y < bounds.getMaxY();
    }

    private static boolean approxEqual(double a, double b) {
        return Math.abs(a - b) <= SCREEN_BOUNDS_MATCH_EPSILON;
    }

    private static double clamp(double v, double lo, double hi) {
        if (hi < lo) {
            return lo;
        }
        return Math.max(lo, Math.min(hi, v));
    }
}
